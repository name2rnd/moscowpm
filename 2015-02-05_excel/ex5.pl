use warnings;
use strict;
use Data::Dumper;
use Spreadsheet::WriteExcel;

sub _decode($) {
  Encode::decode_utf8( shift @_ );
}

my $items = [];
push @$items, { image => 'img/01.jpg', articul => 'АРТ11-РТ',   producer => 'Алмаз', count => 15, price => 57 };
push @$items, { image => 'img/02.jpg', articul => 'АРТ21-РТ',   producer => 'Алмаз', count => 45, price => 69 };
push @$items, { image => 'img/03.jpg', articul => 'АРТ21-ПКС',  producer => 'Алмаз', count => 8,  price => 43 };

my $workbook  = 'Spreadsheet::WriteExcel'->new('ex5.xls');
my $worksheet = $workbook->add_worksheet();
$worksheet->hide_gridlines(2);

# ширина колонок
$worksheet->set_column( 0, 2, 20 );
$worksheet->set_column( 3, 4, 10 );

# header Товары
my $format_header1 = $workbook->add_format( align => 'center', bold => 1, size => 14, font => 'Calibri', 
                                            top => 1, left => 1, right => 1, bottom => 6 );
$worksheet->merge_range( 'A1:C1', _decode('Товары'), $format_header1 );

# header Остатки, Цена
my $format_header2 = $workbook->add_format( align => 'center', valign => 'vcenter', bold => 1, 
                                            size => 14, font => 'Calibri',
                                            top => 1, left => 1, right => 1, bottom => 6 );
$worksheet->merge_range( 'D1:D2', _decode('Остатки'), $format_header2 );
$worksheet->merge_range( 'E1:E2', _decode('Цена'),    $format_header2 );

# header Картинка, Артикул, Производитель
my $format_header3 = $workbook->add_format( align => 'center', bold => 1, size => 14, font => 'Calibri',
                                            top => 1, left => 1, right => 1, bottom => 6 );
$worksheet->write_string('A2', _decode('Картинка'),       $format_header3);
$worksheet->write_string('B2', _decode('Артикул'),        $format_header3 );
$worksheet->write_string('C2', _decode('Производитель'),  $format_header3 );

# форматы
my %common = ( valign => 'vcenter', size => 11, font => 'Calibri', border => 1 );
my $format_articul    = $workbook->add_format( align => 'left',  %common );
my $format_producer   = $workbook->add_format( align => 'center',  %common );
my $format_count_many = $workbook->add_format( align => 'center',  %common, bg_color => 42 );
my $format_count_few  = $workbook->add_format( align => 'center',  %common, bg_color => 26 );
my $format_count_null = $workbook->add_format( align => 'center',  %common, bg_color => 29 );
my $format_price      = $workbook->add_format( align => 'center',  %common, num_format=>'0.00' );
my $format_total      = $workbook->add_format( align => 'center',  bold => 1, bg_color => 15 );

my $row = 2;

# пишем товары
for my $item (@$items) {
  # выставляем высоту строки. По размеру вставленной картинки она не адаптируется
  $worksheet->set_row($row, 100);
  
  # задаем формат ячейки, где будет картинка - не обязательно
  $worksheet->write_blank($row, 0, $format_articul); 
  $worksheet->insert_image($row, 0, $item->{image}, 20, 20);

  $worksheet->write_string($row, 1, _decode($item->{articul}),  $format_articul );
  $worksheet->write_string($row, 2, _decode($item->{producer}), $format_producer );
  # пишем как число, чтобы работала формула в итоговой строке
  $worksheet->write_number($row, 3, _decode($item->{count}),    $item->{count} > 30 
                                                                ? $format_count_many 
                                                                : $item->{count} > 10
                                                                  ? $format_count_few
                                                                  : $format_count_null );
  $worksheet->write_number($row, 4, _decode($item->{price}),    $format_price );
  $row++;
}

# Итого
for ( 0.. 4 ) { $worksheet->write_blank($row, $_, $format_total) };
$worksheet->write_formula($row, 3, '=SUM(D3:D5)', $format_total ); # тут нужно динамически считать адреса ячеек

$row++;
for my $color ( 0 .. 100 ) {
  my $format = $workbook->add_format( bg_color => $color );
  $worksheet->write_string( $row, 1, $color );
  $worksheet->write_string( $row, 2, q(), $format );
  $row++;
}

$workbook->close();

exit(0);
