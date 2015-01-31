use warnings;
use strict;
use Data::Dumper;
use Excel::Template;
use Encode;
#use utf8;

sub _decode($) {
  Encode::decode_utf8( shift @_ );
}
my $filepath = 'ex4.xls';
my $headers  = [map { TITLE => _decode($_) }, (qw/Модель Производитель Цена/)];
my $data     = [];
push @$data, { MODEL => _decode('CDX-555'),         PRODUCER => _decode('Sony'),     PRICE => 650.50 };
push @$data, { MODEL => _decode('АРП-ЦК'),          PRODUCER => _decode('Алмаз'),    PRICE => 1245.58 };
push @$data, { MODEL => _decode('АРТ. 4187АИ'),     PRODUCER => _decode('Котофей'),  PRICE => 200 };
push @$data, { MODEL => _decode(qq{'}.'567543'),    PRODUCER => _decode('Антилопа'), PRICE => 287 };
my $template = Excel::Template->new( filename => 'ex4.xml' );
$template->param( HEADERS => $headers, DATA => $data );
$template->write_file('ex4.xls');

exit(0);
