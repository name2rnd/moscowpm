use warnings;
use strict;
use Data::Dump qw(dd ddx);
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtUnicode;

my $filepath = 'ex1.xls';
my $sheets   = {};

my $oExcel = Spreadsheet::ParseExcel::Workbook->new();
my $oFmtJ  = Spreadsheet::ParseExcel::FmtUnicode->new( Unicode_Map => 'CP1251' );
my $book   = $oExcel->Parse( $filepath, $oFmtJ );

my $i = 0;
for my $sheet ( $book->worksheets() ) {
  my @rows = ();

  my ( $row_min, $row_max ) = $sheet->row_range();
  my ( $col_min, $col_max ) = $sheet->col_range();

  for my $row ( $row_min .. $row_max ) {
    my @data = ();
    for my $col ( $col_min .. $col_max ) {

      my $cell = $sheet->get_cell( $row, $col );
      # если есть ячейка, записываем ее значение
      if ($cell) {
        push( @data, { value => $cell->Value } );
      }
      # иначе записываем пустую строку для сохранения структуры файла
      else {
        push( @data, { value => q() } );
      }
    }
    push( @rows, \@data );
  }
  $sheets->{ $i++ } = \@rows;
}

{
  local *Data::Dump::quote = sub { return qq("$_[0]"); };
  ddx($sheets);
}

