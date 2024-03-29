package Win32::ExcelSimple;
{
  $Win32::ExcelSimple::VERSION = '0.59';
}
use warnings;
use strict;
use Try::Tiny;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;
use Win32::OLE::NLS qw(:LOCALE :DATE);
use Spreadsheet::Read;   #use cr2cell, cell2cr

# ABSTRACT: a wrap of Win32::OLE excel


use Exporter;
our @ISA       = qw( Exporter );
our @EXPORT    = qw( cell2cr cr2cell );

sub new {
	my ($class_name, $file_name) = @_;
		defined $file_name or die "Error: no filename given";
    	-f $file_name or die "Error: filename '$file_name' doesn't exist";
   	 
 	my $Excel = Win32::OLE->new('Excel.Application', sub {$_[0]->Quit;}) or die "Oops, cannot start Excel";
	
    	$Excel->{DisplayAlerts} = 0;
  		$Win32::OLE::Warn = 2;   
  		my $book = $Excel->Workbooks->Open($file_name);
  		my $self = {
		  'excel_handle'   =>  $Excel,
		  'book_handle'    =>  $book,
	  	   };
	bless $self, $class_name;
	return $self;

}

sub open_sheet{
	my $t = $_[0]->{'book_handle'}->Worksheets($_[1]);
	bless \$t, 'Win32::ExcelSimple::Sheet';
}

sub save_excel{

	my $self = shift;
	$self->{ 'book_handle' }->save();
	return 0;

} 
sub saveas_excel{
	my $self = shift;
	my $name = shift;
	$self->{ 'book_handle' }->saveas($name);
	return 0;
}

sub close_excel{
	my $self = shift;
 	$self->{ 'excel_handle' }->Workbooks->close;

}


sub DESTROY{

	my $self = shift;
	$self->close_excel();
#	print "save all and exit!!!\n";

}
package Win32::ExcelSimple::Sheet;
{
  $Win32::ExcelSimple::Sheet::VERSION = '0.59';
}
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;
use Win32::OLE::NLS qw(:LOCALE :DATE);
sub read{
	my ($sheet_h, $x1,$y1, $x2, $y2) = @_;
	return undef unless(defined($x1) or defined($x2) or defined($y1) or defined($y2));  
	return undef unless($x1 =~ /^\d+$/ or $y1 =~ /^\d+$/ or $x2 =~ /^\d+$/);
	if ($x1 == $x2 and $y1 == $y2) { 
		return $$sheet_h->Cells($y1, $x1)->{Value};
	}
	my $address = Win32::ExcelSimple::cr2cell($x1,$y1) . ':' . Win32::ExcelSimple::cr2cell($x2, $y2);
    my $data = $$sheet_h->Range($address)->{Value};
	if ($x1 == $x2){
	my @a =   map {@{$_}} @{$data};
    return \@a;	
	}
	elsif ($y1 == $y2){
		   return @{$data};
	}
	else{
		return $data;
	}
}


sub get_last_row{
	my $sheet_h = shift;
    return ${$sheet_h}->UsedRange->Find({What=>"*",
    			SearchDirection=> xlPrevious,
    			SearchOrder=> xlByRows})->{Row};

}

sub get_last_col{
	my $sheet_h = shift;
	return ${$sheet_h}->UsedRange->Find({What=>"*", 
                  SearchDirection=> xlPrevious,
                  SearchOrder=> xlByColumns})->{Column};
	  }

sub read_cell{
	my ($sheet_h, $x, $y) = @_;
	return undef unless(defined($x) or defined($x));  
	  return undef unless $x =~ /^\d+$/;
	  return undef unless $y =~ /^\d+$/;
	return $$sheet_h->Cells($y, $x)->{Value};
}
sub write_cell{
    my ($sheet_h, $x, $y, $data) = @_;
	my $address = Win32::ExcelSimple::cr2cell($x,$y);
	return undef unless $address;
	$data = [] unless defined $data;
	return ${$sheet_h}->Range($address)->{Value} = $data;

}
sub write_row{
	my ($sheet_h, $x1,$y1, $data) = @_;
	return $sheet_h->write_cell($x1, $y1, $data) if ref $data ne ref [];
	return $sheet_h->write_cell($x1, $y1, undef)    unless @$data;
	my  $x2 = $x1+ $#{$data};
	my  $y2 = $y1;
	my $address = Win32::ExcelSimple::cr2cell($x1,$y1) . ':' . Win32::ExcelSimple::cr2cell($x2, $y2);
	return undef unless $address;
	$$sheet_h->Range($address)->{Value} = [$data];
}

sub write_col{
	my ($sheet_h, $x1,$y1, $data) = @_;
	return $sheet_h->write_cell($x1, $y1, $data) if ref $data ne ref [];
	for (@{$data}){
	$sheet_h->write_row($x1, $y1, $_);
	$y1++;
}
}

sub write{
	my ($sheet_h, $x, $y, $data) = @_;
	         if( ref ${$data}[0] ne ref []){
				 $sheet_h->write_row($x,$y,$data);
			 }
			 else{
				 for(my $i = 0; $i < (scalar @{$data}); $i ++){
						 $sheet_h->write_row($x, $y, $data->[$i]);
							 $y ++;
				 }
             }
}

sub cell_walk{
    my ($sheet_h, $x1, $y1, $x2, $y2, $callback, $callback_data) = @_;
	for ( my $row = $x1 ; $row <= $x2 ; $row++ ) {
    	for ( my $col = $y1 ; $col <= $y2 ; $col++ ) {
			  $callback->($sheet_h->Cells( $row, $col ), $callback_data);
		}
	}

}

sub whole_walk{
    my ($self, $callback) = @_;
	my $x = [1,1];

my $y = [$self->get_last_row(), $self->get_last_col()];

$self->cell_walk($x, $y, $callback);

}

1; # End of Win32::ExcelSimple

__END__

=pod

=head1 NAME

Win32::ExcelSimple - a wrap of Win32::OLE excel

=head1 VERSION

version 0.59

=head1 SYNOPSIS

Quick summary of what the module does.

Perhaps a little code snippet.

    use Win32::ExcelSimple;

    my $foo = Win32::ExcelSimple->new();
    ...
	see test files for details

=head1 DESCRIPTION

Win32::ExcelSimple is a thin wrap of Win32::OLE Excel. The behavior is much like SpreadSheet::Write but with ability of modifying existing Excel file.
Note: this module is based on CELL address. You need to use cr2cell or cell2cr to translate number address. 

=head1 NAME

Win32::ExcelSimple -  a easier way to use Microsoft Excel simplier

=head1 VERSION

Version 0.59

=head1 AUTHOR

Andy Xiao, C<< <andy.xiao at cpan.org> >>

=head1 BUGS

Please report any bugs or feature requests to C<bug-win32-excelsimple at rt.cpan.org>, or through
the web interface at L<http://rt.cpan.org/NoAuth/ReportBug.html?Queue=Win32-ExcelSimple>.  I will be notified, and then you'll
automatically be notified of progress on your bug as I make changes.

=head1 SUPPORT

You can find documentation for this module with the perldoc command.

    perldoc Win32::ExcelSimple

You can also look for information at:

=over 4

=item * RT: CPAN's request tracker (report bugs here)

L<http://rt.cpan.org/NoAuth/Bugs.html?Dist=Win32-ExcelSimple>

=item * AnnoCPAN: Annotated CPAN documentation

L<http://annocpan.org/dist/Win32-ExcelSimple>

=item * CPAN Ratings

L<http://cpanratings.perl.org/d/Win32-ExcelSimple>

=item * Search CPAN

L<http://search.cpan.org/dist/Win32-ExcelSimple/>

=back

=head1 LICENSE AND COPYRIGHT

Copyright 2011 2012 Andy Xiao.

This program is free software; you can redistribute it and/or modify it
under the terms of either: the GNU General Public License as published
by the Free Software Foundation; or the Artistic License.

See http://dev.perl.org/licenses/ for more information.

=head1 AUTHOR

xiaoyafeng <xyf.xiao@gmail.com>

=head1 COPYRIGHT AND LICENSE

This software is copyright (c) 2012 by xiaoyafeng.

This is free software; you can redistribute it and/or modify it under
the same terms as the Perl 5 programming language system itself.

=cut
