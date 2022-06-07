package Excel::ValueWriter::XLSX;
use 5.024;
use strict;
use warnings;
use utf8;
use Archive::Zip          qw/AZ_OK/;
use Scalar::Util          qw/looks_like_number/;
use List::Util            qw/none max any/;
use Params::Validate      qw/validate_with SCALAR SCALARREF ARRAYREF UNDEF/;
use POSIX                 qw/strftime/;
use Date::Calc            qw/Delta_Days/;
use Carp                  qw/croak/;
use Encode                qw/encode_utf8/;

my $VERSION = '0.2';


# TODO
# - handle 1904



#======================================================================
# GLOBALS
#======================================================================

my $DATE_STYLE = 1;                        # 0-based index into the <cellXfs> format for dates ..
                                           # .. defined in the styles() method

my $SHEET_NAME = qr(^[^\\/?*\[\]]{1,31}$); # valid sheet names: <32 chars, no chars \/?*[] 
my $TABLE_NAME = qr(^\w{3,}$);             # valid table names: >= 3 chars, no spaces


# specification in Params::Validate format for checking parameters to the new() method 
my %params_spec = (

  # date_regex : for identifying dates in data cells. Should capture into $+{d}, $+{m} and $+{y}.
  date_regex => {type => SCALARREF|UNDEF, optional => 1, default =>
                  qr[^(?: (?<d>\d\d?)    \. (?<m>\d\d?) \. (?<y>\d\d\d\d)  # dd.mm.yyyy
                        | (?<y>\d\d\d\d) -  (?<m>\d\d?) -  (?<d>\d\d?)     # yyyy-mm-dd
                        | (?<m>\d\d?)    /  (?<d>\d\d?) /  (?<y>\d\d\d\d)) # mm/dd/yyyy
                      $]x},

  template         => {type => SCALAR,   optional => 1},
  sheets_to_remove => {type => ARRAYREF, optional => 1},

 );


#======================================================================
# CONSTRUCTOR
#======================================================================

sub new {
  my $class = shift;

  # check parameters and create $self
  my $self = validate_with( params      => \@_,
                            spec        => \%params_spec,
                            allow_extra => 0,
                           );

  # initial values for internal data structures
  $self->{sheets}                = {}; # ($sheet_name => $sheet_index)
  $self->{tables}                = {}; # ($table_name => $table_index)
  $self->{shared_strings}        = {}; # ($string => $string_index)
  $self->{n_strings_in_workbook} = 0;  # total nb of strings (including duplicates)
  $self->{last_string_id}        = 0;  # index for the next shared string

  # immediately open a Zip archive
  $self->{zip} = Archive::Zip->new;

  bless $self, $class;

  $self->load_template if $self->{template};

  return $self;
}


#======================================================================
# GATHERING DATA
#======================================================================


sub add_sheet {
  # 3rd parameter ($headers) may be omitted -- so we insert an undef if necessary
  splice @_, 3, 0, undef if @_ < 5;

  # now we can parse the parameters
  my ($self, $sheet_name, $table_name, $headers, $code_or_array) = @_;

  # check if the given sheet name is valid
  $sheet_name =~ $SHEET_NAME
    or croak "'$sheet_name' is not a valid sheet name";

  # register the sheet name 
  my $sheet = $self->{sheets}{$sheet_name};
  if ($sheet) {
    $sheet->{to_remove}
      or croak "'this workbook already has a sheet named '$sheet_name'";
  }
  else {
    $sheet = $self->{sheets}{$sheet_name} = {id => $self->max_sheet_id + 1};
  }


  # local copy for convenience
  my $date_regex = $self->{date_regex};

  # iterator for generating rows; either received as argument or built as a closure upon an array
  my $next_row 
    = ref $code_or_array eq 'CODE'  ? $code_or_array
    : ref $code_or_array ne 'ARRAY' ? croak 'add_sheet() : missing or invalid $rows argument'
    : do {my $i = 0; sub { $i < @$code_or_array ? $code_or_array->[$i++] : undef}};

  # if $headers were not given explicitly, the first row will do
  $headers //= $next_row->();

  # array of column references in A1 Excel notation
  my @col_letters = ('A'); # this array will be expanded on demand in the loop below


  # start building XML for the sheet
  my @xml = (
    q{<?xml version="1.0" encoding="UTF-8" standalone="yes"?>},
    q{<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"},
              q{ xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">},
    q{<sheetData>},
    );

  # loop over rows and columns
  my $row_num = 0;
 ROW:
  for (my $row = $headers; $row; $row = $next_row->()) {
    $row_num++;
    my $last_col = @$row or next ROW;
    my @cells;

  COLUMN:
    foreach my $col (0 .. $last_col-1) {

      # if this column letter is not known yet, compute it using Perl's increment op on strings
      my $col_letter = $col_letters[$col]
                   //= do {my $prev_letter = $col_letters[$col-1]; ++$prev_letter};

      # get the value; if the cell is empty, no need to write it into the XML
      my $val = $row->[$col];
      defined $val and length $val or next COLUMN;

      # choose XML attributes and inner value
      (my $attrs, $val)
        = looks_like_number $val             ? (""                  , $val                          )
        : $date_regex && $val =~ $date_regex ? (qq{ s="$DATE_STYLE"}, n_days($+{y}, $+{m}, $+{d})   )
        :                                      (qq{ t="s"}          , $self->add_shared_string($val));

      # add the new XML cell
      push @cells, sprintf qq{<c r="%s%d"%s><v>%s</v></c>}, $col_letter, $row_num, $attrs, $val;
    }

    # generate the row XML and add it to the sheet
    my $row_xml = join "", qq{<row r="$row_num" spans="1:$last_col">}, @cells, qq{</row>};
    push @xml, $row_xml;
  }

  # close sheet data
  push @xml, q{</sheetData>};

  # if required, add a table corresponding to this sheet into the zip archive, and refer to it in XML
  if ($table_name && $row_num) {
    my $table_id = $self->add_table($table_name, $col_letters[-1], $row_num, @$headers);
    push $sheet->{table_ids}->@*, $table_id;
    push @xml, q{<tableParts count="1"><tablePart r:id="rId1"/></tableParts>};
  }

  # close the worksheet xml
  push @xml, q{</worksheet>};

  # insert the sheet XML and its rels into the zip archive
  my $sheet_file = "sheet$sheet->{id}.xml";
  $self->{zip}->addString(join("", @xml), "xl/worksheets/$sheet_file");
  $self->{zip}->addString($self->worksheet_rels($sheet), "xl/worksheets/_rels/$sheet_file.rels");

  delete $sheet->{to_remove};

  return $sheet->{id};
}


sub max_sheet_id {
  my ($self) = @_;

  return max 0, map {$_->{id}} values $self->{sheets}->%*;
}

sub max_table_id {
  my ($self) = @_;

  return max 0, map {$_->{id}} values $self->{tables}->%*;
}


sub add_shared_string {
  my ($self, $string) = @_;

  # keep a global count of how many strings are in the workbook
  $self->{n_strings_in_workbook}++;

  # if that string was already stored, return its id, otherwise create a new id
  $self->{shared_strings}{$string} //= $self->{last_string_id}++;
}



sub add_table {
  my ($self, $table_name, $last_col, $last_row, @col_names) = @_;

  # check if the given table name is valid
  $table_name =~ $TABLE_NAME
    or croak "'$table_name' is not a valid table name";

  # register the table name 
  ! $self->{tables}{$table_name}
    or croak "'this workbook already has a table named '$table_name'";
  my $table = $self->{tables}{$table_name} 
            = {id => $self->max_table_id + 1};  # THINK: other fields in this subhash ?

  # build column headers from first data row
  unshift @col_names, undef; # so that the first index is at 1, not 0
  my @columns = map {qq{<tableColumn id="$_" name="$col_names[$_]"/>}} 1 .. $#col_names;

  # Excel range of this table
  my $ref = "A1:$last_col$last_row";

  # assemble XML for the table
  my @xml = (
    qq{<?xml version="1.0" encoding="UTF-8" standalone="yes"?>},
    qq{<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"}.
         qq{ id="$table->{id}" displayName="$table_name" ref="$ref" totalsRowShown="0">},
    qq{<autoFilter ref="$ref"/>},
    qq{<tableColumns count="$#col_names">},
    @columns,
    qq{</tableColumns>},
    qq{<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>},
    qq{</table>},
   );

  # insert into the zip archive
  $self->{zip}->addString(encode_utf8(join "", @xml), "xl/tables/table$table->{id}.xml");

  return $table->{id};
}


sub worksheet_rels {
  my ($self, $sheet) = @_;

  my @rels = map {("officeDocument/2006/relationships/table" => "../tables/table$_.xml")}
                 $sheet->{table_ids}->@*;

  return $self->relationships(@rels);
}


#======================================================================
# BUILDING THE ZIP CONTENTS
#======================================================================

sub save_as {
  my ($self, $file_name) = @_;

  $self->remove_unused_sheets;
  $self->recompute_sheet_indices;
  $self->recompute_table_indices;

  # assemble all parts within the zip, except sheets and tables that were already added previously
  my $zip = $self->{zip};
  $zip->addString($self->content_types,      "[Content_Types].xml");
  $zip->addString($self->core,               "docProps/core.xml");
  $zip->addString($self->app,                "docProps/app.xml");
  $zip->addString($self->workbook,           "xl/workbook.xml");
  $zip->addString($self->_rels,              "_rels/.rels");
  $zip->addString($self->workbook_rels,      "xl/_rels/workbook.xml.rels");
  $zip->addString($self->shared_strings,     "xl/sharedStrings.xml");
  $zip->addString($self->styles,             "xl/styles.xml");

  # write the Zip archive
  my $write_result = $zip->writeToFileNamed($file_name);
  $write_result == AZ_OK
    or croak "could not write into $self->{xlsx}";
}


sub remove_unused_sheets {
  my ($self) = @_;

  while (my ($name, $sheet) = each $self->{sheets}->%*) {
    delete $self->{sheets}{$name} if $sheet->{to_remove};
  }
}


sub recompute_sheet_indices {
  my ($self) = @_;

  my $next_sheet_id = 1;
  foreach my $sheet (sort {$a->{id} <=> $b->{id}} values $self->{sheets}->%*) {
    $sheet->{remap} = $next_sheet_id if $sheet->{id} != $next_sheet_id;
    $next_sheet_id++;
  }

  foreach my $sheet (grep {$_->{remap}} values $self->{sheets}->%*) {

    warn "REMAP sheet $sheet->{id} TO $sheet->{remap}\n";

    my $old_name = $self->sheet_member($sheet->{id});
    my $new_name = $self->sheet_member($sheet->{remap});
    $self->rename_zip_member($old_name, $new_name);

    # do the same for sheet .rels
    s[(sheet\d+\.xml)$][_rels/$1.rels] for $old_name, $new_name;
    $self->rename_zip_member($old_name, $new_name);

    $sheet->{id} = delete $sheet->{remap};
  }
}





sub recompute_table_indices {
  my ($self) = @_;

  my %remap;


  my $next_table_id = 1;
  foreach my $table (sort {$a->{id} <=> $b->{id}} values $self->{tables}->%*) {
    $remap{$table->{id}} = $next_table_id if $table->{id} != $next_table_id;
    $next_table_id++;
  }

  foreach my $table (grep {$remap{$_->{id}}} values $self->{tables}->%*) {
    my $new_id = $remap{$table->{id}};

    warn "REMAP TABLE  $table->{id} TO $new_id\n";

    my $old_name = $self->table_member($table->{id});
    my $new_name = $self->table_member($new_id);
    $self->rename_zip_member($old_name, $new_name);

    $table->{id} = $new_id;
  }

  # recompute sheet rels if table ids have changed
  foreach my $sheet (values $self->{sheets}->%*) {
    if (any {$remap{$_}} $sheet->{table_ids}->@*) {
      my @new_table_ids = map {$remap{$_} // $_} $sheet->{table_ids}->@*;
      $sheet->{table_ids} = \@new_table_ids;
      my $rels_file = "xl/worksheets/_rels/sheet$sheet->{id}.xml.rels";
      $self->zip->removeMember($rels_file);
      $self->{zip}->addString($self->worksheet_rels($sheet), $rels_file);
    }
  }
}



sub rename_zip_member {
  my ($self, $old_name, $new_name) = @_;

  my $member = $self->zip->removeMember($old_name);
  $member->fileName($new_name);
  $self->zip->addMember($member);
}


sub sheet_member {
  my ($self, $sheet_id) = @_;
  return "xl/worksheets/sheet$sheet_id.xml";
}


sub table_member {
  my ($self, $table_id) = @_;
  return "xl/tables/table$table_id.xml";
}


sub _rels {
  my ($self) = @_;

  return $self->relationships("officeDocument/2006/relationships/extended-properties" => "docProps/app.xml",
                              "package/2006/relationships/metadata/core-properties"   => "docProps/core.xml",
                              "officeDocument/2006/relationships/officeDocument"      => "xl/workbook.xml");
}

sub workbook_rels {
  my ($self) = @_;

  my @rels = map {("officeDocument/2006/relationships/worksheet"     => "worksheets/sheet$_->{id}.xml")}
                 sort {$a->{id} <=> $b->{id}} values $self->{sheets}->%*;
  push @rels,      "officeDocument/2006/relationships/sharedStrings" => "sharedStrings.xml",
                   "officeDocument/2006/relationships/styles"        => "styles.xml";

  return $self->relationships(@rels);
}


sub workbook {
  my ($self) = @_;

  # opening XML
  my @xml = (
    qq{<?xml version="1.0" encoding="UTF-8" standalone="yes"?>},
    qq{<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"},
             qq{ xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">},
    qq{<sheets>},
    );

  # references to the worksheets
  my @ordered_sheet_names = sort {$self->{sheets}{$a}{id} <=> $self->{sheets}{$b}{id}} keys $self->{sheets}->%*;
  my $rId = 1;
  foreach my $sheet_name (@ordered_sheet_names) {
    push @xml, qq{<sheet name="$sheet_name" sheetId="$rId" r:id="rId$rId"/>};
    $rId++;  # THINK : not clear if rId must absolutely be the same as $sheet->{id}
  }

  # closing XML
  push @xml, q{</sheets>}, q{</workbook>};

  return encode_utf8(join "", @xml);
}





sub content_types {
  my ($self) = @_;

  my $spreadsheetml = "application/vnd.openxmlformats-officedocument.spreadsheetml";

  my @sheets_xml
    = map {qq{<Override PartName="/xl/worksheets/sheet$_->{id}.xml" ContentType="$spreadsheetml.worksheet+xml"/>}} sort {$a->{id} <=> $b->{id}} values $self->{sheets}->%*;

  my @tables_xml
    = map {qq{  <Override PartName="/xl/tables/table$_->{id}.xml" ContentType="$spreadsheetml.table+xml"/>}} sort {$a->{id} <=> $b->{id}} values $self->{tables}->%*;

  my @xml = (
    qq{<?xml version="1.0" encoding="UTF-8" standalone="yes"?>},
    qq{<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">},
    qq{<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>},
    qq{<Default Extension="xml" ContentType="application/xml"/>},
    qq{<Override PartName="/xl/workbook.xml" ContentType="$spreadsheetml.sheet.main+xml"/>},
    qq{<Override PartName="/xl/styles.xml" ContentType="$spreadsheetml.styles+xml"/>},
    qq{<Override PartName="/xl/sharedStrings.xml" ContentType="$spreadsheetml.sharedStrings+xml"/>},
    qq{<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>},
    qq{<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>},
    @sheets_xml,
    @tables_xml,
    qq{</Types>},
   );

  return join "", @xml;
}


sub core {
  my ($self) = @_;

  my $now = strftime "%Y-%m-%dT%H:%M:%SZ", gmtime;

  my @xml = (
    qq{<?xml version="1.0" encoding="UTF-8" standalone="yes"?>},
    qq{<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" },
                      qq{ xmlns:dc="http://purl.org/dc/elements/1.1/"},
                      qq{ xmlns:dcterms="http://purl.org/dc/terms/"},
                      qq{ xmlns:dcmitype="http://purl.org/dc/dcmitype/"},
                      qq{ xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">},
    qq{<dcterms:created xsi:type="dcterms:W3CDTF">$now</dcterms:created>},
    qq{<dcterms:modified xsi:type="dcterms:W3CDTF">$now</dcterms:modified>},
    qq{</cp:coreProperties>},
   );

  return join "", @xml;
}

sub app {
  my ($self) = @_;

  my @xml = (
    qq{<?xml version="1.0" encoding="UTF-8" standalone="yes"?>},
    qq{<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"},
               qq{ xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">},
    qq{<Application>Microsoft Excel</Application>},
    qq{</Properties>},
   );

  return join "", @xml;
}




sub shared_strings {
  my ($self) = @_;

  # array of XML nodes for each shared string
  my @si_nodes;
  while (my ($string, $index) = each $self->{shared_strings}->%*) {
    $si_nodes[$index] = si_node($string);
  }

  # assemble XML
  my @xml = (
    qq{<?xml version="1.0" encoding="UTF-8" standalone="yes"?>},
    qq{<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"},
         qq{ count="$self->{n_strings_in_workbook}" uniqueCount="$self->{last_string_id}">},
    @si_nodes,
    qq{</sst>},
   );

  return encode_utf8(join "", @xml);
}


sub styles {
  my ($self) = @_;

  # minimal stylesheet
  # style "1" will be used for displaying dates; it uses the default numFmtId for dates, which is 14 (Excel builtin).
  # other nodes are empty but must be present
  my @xml = (
    q{<?xml version="1.0" encoding="UTF-8" standalone="yes"?>},
    q{<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">},
    q{<fonts count="1"><font/></fonts>},
    q{<fills count="1"><fill/></fills>},
    q{<borders count="1"><border/></borders>},
    q{<cellStyleXfs count="1"><xf/></cellStyleXfs>},
    q{<cellXfs count="2"><xf/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs>},
    q{<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>},
    q{</styleSheet>},
   );

  my $xml = join "", @xml;

  return $xml;
}


#======================================================================
# UTILITY METHODS
#======================================================================

sub relationships {
  my ($self, @rels) = @_;

  # build a "rel" file from a list of relationships
  my @xml = (
    qq{<?xml version="1.0" encoding="UTF-8" standalone="yes"?>},
    qq{<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">},
   );

  my $id = 1;
  while (my ($type, $target) = splice(@rels, 0, 2)) {
    push @xml, qq{<Relationship Id="rId$id" Type="http://schemas.openxmlformats.org/$type" Target="$target"/>};
    $id++;
  }

  push @xml, qq{</Relationships>};

  return join "", @xml;
}




#======================================================================
# UTILITY ROUTINES
#======================================================================


my %entity = ( '<' => '&lt;', '>' => '&gt;', '&' => '&amp;' );

sub si_node {
  my ($string) = @_;

  # build XML node for a single shared string
  $string =~ s/([<>&])/$entity{$1}/eg;
  my $maybe_preserve_space = $string =~ /^\s|\s$/ ? ' xml:space="preserve"' : '';
  my $node = qq{<si><t$maybe_preserve_space>$string</t></si>};

  return $node;
}

sub n_days {
  my ($y, $m, $d) = @_;

  # convert the given date into a number of days since 1st January 1900
  my $n_days = Delta_Days(1900, 1, 1, $y, $m, $d) + 1;
  my $is_after_february_1900 = $n_days > 59;
  $n_days += 1 if $is_after_february_1900; # because Excel wrongly treats 1900 as a leap year

  return $n_days;
}






#======================================================================
# LOAD AN EXISTING EXCEL FILE AS A TEMPLATE
#======================================================================

my $rx_string_in_cell = qr[<c([^>]+?)t="s"><v>(\d+)</v></c>];

sub load_template {
  my ($self) = @_;

  # load ZIP archive
  my $read_result = $self->{zip}->read($self->{template});
  $read_result == AZ_OK  or die "cannot unzip $self->{template}";


  # extract sheet names
  my $workbook_xml = $self->_zip_member_contents('xl/workbook.xml');
  my @sheet_names = ($workbook_xml =~ m[<sheet name="(.+?)"]g);
  $self->{sheets}{$sheet_names[$_]}{id} = $_ + 1 for 0 .. $#sheet_names;

  # mark sheets to remove
  foreach my $sheet_name ($self->{sheets_to_remove}->@*) {
    my $sheet =   $self->{sheets}{$sheet_name}
      or die "can't remove sheet '$sheet_name' : absent in template";
    $sheet->{to_remove} = 1;
  }

  # parse existing sheets to gather string indices and table indices
  my @keep_string;
  $self->parse_sheet($_, \@keep_string) foreach values $self->{sheets}->%*;

  # recompute indices for strings that must be kept
  my $next_string_index = 0;
  my @remap_string;
  $keep_string[$_] and $remap_string[$_] = $next_string_index++ for 0 .. $#keep_string;

  # adapt string indices in XML of sheets that will be kept
  foreach my $sheet (grep {!$_->{to_remove}} values $self->{sheets}->%*) {
    $sheet->{xml} =~ s[$rx_string_in_cell]
                      [<c$1t="s"><v>$remap_string[$2]</v></c>]g;

    $self->zip->addString(delete $sheet->{xml},
                          "xl/worksheets/sheet$sheet->{id}.xml");

  }

  # feed initial structure for shared strings
  $self->init_shared_strings(\@remap_string);


  # members that will be automatically regenerated must first be removed from zip
  $self->{zip}->removeMember($_) for ("[Content_Types].xml",
                                      "docProps/core.xml",
                                      "docProps/app.xml",
                                      "xl/workbook.xml",
                                      "_rels/.rels",
                                      "xl/_rels/workbook.xml.rels",
                                      "xl/sharedStrings.xml",
                                      "xl/styles.xml",
                                      'xl/sharedStrings.xml',
                                     );
}




sub parse_sheet {
  my ($self, $sheet, $keep_string) = @_;

  # extract sheet XML
  my $sheet_file = "sheet$sheet->{id}.xml";
  my $zip_member = $self->zip->removeMember("xl/worksheets/$sheet_file");
  my $sheet_xml  = $zip_member->contents;
  utf8::decode($sheet_xml);


  # if the sheet is not to be removed, gather all string ids in that sheet
  if (!$sheet->{to_remove}) {
    while ($sheet_xml =~ m/$rx_string_in_cell/g) {
      my $string_id = $2;
      $keep_string->[$string_id]++;
      $self->{n_strings_in_workbook}++;
    }
    # store the XML because we need to rewrite it later
    $sheet->{xml} = $sheet_xml;
  }

  # extract or remove sheet rels and register ids of tables in that sheet
  my $rels_filename = "xl/worksheets/_rels/$sheet_file.rels";
  my $meth          = $sheet->{to_remove} ? 'removeMember' : 'memberNamed';
  my $rels_member   = $self->zip->$meth($rels_filename);
  my $rels_xml      = $rels_member->contents;
  $sheet->{table_ids}
    = [($rels_xml =~ m[relationships/table" Target="../tables/table(\d+).xml"]g)];

  foreach my $table_id ($sheet->{table_ids}->@*) {
    my $member_name = "xl/tables/table$table_id\.xml";
    if ($sheet->{to_remove}) {
      $self->zip->removeMember($member_name);
    }
    else {
      my $table_xml = $self->_zip_member_contents($member_name);
      my ($table_name) = ($table_xml =~ m{<table.+?displayName="(\w+)"});
      $self->{tables}{$table_name} = {id => $table_id};
    }
  }
}



sub init_shared_strings {
  my ($self, $remap_index) = @_;

  my $strings_xml = $self->_zip_member_contents('xl/sharedStrings.xml');
  my $old_string_index = 0;
  my @strings_to_keep;
  while ($strings_xml =~ m[<si>(.*?)</si>]sg) {
    my $innerXML  = $1;
    my $new_index = $remap_index->[$old_string_index];
    if (defined $new_index) {
      # concatenate contents from all <t> nodes (usually there is only 1) and decode XML entities
      my $string = join "", ($innerXML =~ m[<t[^>]*>(.+?)</t>]sg);
      _decode_xml_entities($string);
      $self->{shared_strings}{$string} = $new_index;
    }
    $old_string_index++;
  }

  $self->{last_string_id} = scalar keys $self->{shared_strings}->%*;
}



sub zip {shift->{zip}}

  ## COPIED FROM VALUEREADER !!
  

sub _zip_member_contents {
  my ($self, $member) = @_;

  my $contents = $self->zip->contents($member)
    or die "no contents for member $member";
  utf8::decode($contents);

  return $contents;
}


sub _zip_member_name_for_sheet {
  my ($self, $sheet_name) = @_;

  # check that sheet name was given
  $sheet_name or die "missing sheet name";

  # get sheet id
  my $id = $self->{workbook_data}{sheets}{$sheet_name}
    or die "no such sheet: $sheet_name";

  # construct member name for that sheet
  return "xl/worksheets/sheet$id.xml";
}




sub _decode_xml_entities {
  state $xml_entities   = { amp  => '&',
                            lt   => '<',
                            gt   => '>',
                            quot => '"',
                            apos => "'",
                           };
  state $entity_names   = join '|', keys %$xml_entities;
  state $regex_entities = qr/&($entity_names);/;

  # substitute in-place
  $_[0] =~ s/$regex_entities/$xml_entities->{$1}/eg;
}



1;
__END__

=encoding utf-8

=head1 NAME

Excel::ValueWriter::XLSX - generating data-only Excel workbooks in XLSX format, fast

=head1 SYNOPSIS

  my $writer = Excel::ValueWriter::XLSX->new;
  $writer->add_sheet($sheet_name1, $table_name1, [qw/a b/], [[1, 2], [3, 4]]);
  $writer->add_sheet($sheet_name2, $table_name2, \@headers, $row_generator);
  $writer->save_as($filename);


=head1 DESCRIPTION

The common way for generating Microsoft Excel workbooks in C<XLSX>
format from Perl programs is the excellent L<Excel::Writer::XLSX>
module. That module is very rich in features, but quite costly in CPU
and memory usage. By contrast, the present module
L<Excel::ValueWriter::XLSX> is aimed at fast and cost-effective
production of data-only workbooks, containing nothing but plain
values. Such workbooks are useful in architectures where Excel is used
merely as a local database, for example in connection with a PowerBI
architecture.

=head1 VERSION

This is version 0.1, the first release.
Until version 1.0, slight changes may occur in the API.


=head1 METHODS

=head2 new

  my $writer = Excel::ValueWriter::XLSX->new(%options);

Constructor for a new writer object. Currently the only option is :

=over

=item date_regex

A compiled regular expression for detecting data cells that contain dates.
The default implementation recognizes dates in C<dd.mm.yyyy>, C<yyyy-mm-dd>
and C<mm/dd/yyyy> formats. User-supplied regular expressions should use
named captures so that the day, month and year values can be found respectively
in C<< $+{d} >>, C<< $+{m} >> and C<< $+{y} >>.

=back

=head2 add_sheet

  $writer->add_sheet($sheet_name, $table_name, [$headers,] $rows);

Adds a new worksheet into the workbook.

=over

=item *

The C<$sheet_name> is mandatory; it must be unique and between 1 and 31 characters long.

=item *

The C<$table_name> is optional; if not C<undef>, the sheet contents
will be registered as an Excel table. The table name must be unique,
of minimum 3 characters, without spaces or special characters.

=item *

The C<$headers> argument is optional; it may be C<undef> or may even be absent.
If present, it should contain an arrayref of scalar values, that will
be used as column names for the table associated with that worksheet.
Column names should be unique (otherwise Excel will automatically add
a discriminating number). If C<$headers> are not present, the first
row in C<$rows> will be treated as headers.


=item *

The C<$rows> argument may be either a reference to a 2-dimensional array of values,
or a reference to a callback function that will return a new row at each call, in the
form of a 1-dimensional array reference. An empty return from the callback
function signals the end of data (but intermediate empty rows may be returned
as C<< [] >>). Callback functions should typically be I<closures> over a lexical
variable that remembers when the last row has been met. Here is an example of a
callback function used to feed a sheet with 500 lines of 300 columns of random numbers:

  my @headers_for_rand = map {"h$_"} 1 .. 300;
  my $random_rows = do {my $count = 500; sub {$count-- > 0 ? [map {rand()} 1 .. 300] : undef}};
  $writer->add_sheet(RAND_SHEET => rand => \@headers_for_rand, $random_rows);

=back

Cells within a row must contain scalar values. Values that look like numbers are treated
as numbers, string values that match the C<date_regex> are converted into numbers and
displayed through a date format, all other strings are treated as shared strings at the
workbook level (hence a string that appears several times in the input data will be stored
only once within the workbook).

=head2 save_as

  $writer->save_as($filename);

Writes the workbook contents into the specified C<$filename>.


=head1 ARCHITECTURAL NOTE

Âlthough I'm a big fan of L<Moose> and its variants, the present module is implemented
in POPO (Plain Old Perl Object) : since the aim is to maximize cost-effectiveness, and since
the object model is extremely simple, there was no ground for using a sophisticated object system.

=head1 SEE ALSO

L<Excel::Writer::XLSX>

=head1 BENCHMARKS

Not done yet

=head1 TO DO

  - options for workbook properties : author, etc.
  - support for 1904 date schema


=head1 AUTHOR

Laurent Dami, E<lt>dami at cpan.orgE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright 2022 by Laurent Dami.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.



    foreach my $table_id (@table_ids) {
      my $table_xml = $self->_zip_member_contents("xl/tables/table$table_id.xml");
      my ($table_name) = $table_xml =~ m[id="\d+" displayName="(\w+)"];
      $self->{tables}{$table_name} = $table_id;

      # THINK : should also parse tables in sheets that will be removed

    }



  # recompute indices for tables that must be kept
  my $next_table_index = 0;
  my @remap_table;
  my @table_ids = sort  {$a <=> $b} keys %table_by_id;
  !$table_by_id{$_}{to_remove} and $remap_table[$_] = $next_table_index++ for @table_ids;




TABLE INFO
  - worksheet XML : static link to worksheet rels
  - worksheet rels : link to table.xml
  - content-type

