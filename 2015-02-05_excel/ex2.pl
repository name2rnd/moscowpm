use warnings;
use strict;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::FmtUnicode;

my $filepath = 'ex2.xls';

my $oExcel = Spreadsheet::ParseExcel::Workbook->new();
my $oFmtJ  = Spreadsheet::ParseExcel::FmtUnicode->new( Unicode_Map => 'CP1251' );
my $book    = $oExcel->Parse( $filepath, $oFmtJ );

for my $sheet ( $book->worksheets() ) {

  my ( $row_min, $row_max ) = $sheet->row_range();
  my ( $col_min, $col_max ) = $sheet->col_range();

  for my $row ( $row_min .. $row_max ) {
    print '='x10, $row, '='x10, "\n";
    for my $col ( $col_min .. $col_max ) {
      my $cell = $sheet->get_cell( $row, $col );

      my $format = $cell->get_format();
      my $font   = $format->{ Font };
      
      printf "[%i:%i]\n",                   $row, $col;
      printf "val           = %s\n",        $cell->value();
      printf "val_u         = %s\n",        $cell->unformatted();
      printf "font_size     = %i\n",        $font->{ Height };
      printf "font_colorRGB = %s\n",        Spreadsheet::ParseExcel->ColorIdxToRGB( $font->{ Color } );
      printf "color_fill    = %i %i\n",     $format->{ Fill }->[1], $format->{ Fill }->[2];
      printf "level         = %i\n",        $format->{ Indent }; 
      print "\n";
    }
  }
}
