#!/usr/bin/perl

# A script to beautify automatic log output from Moseley CommServer by
# presenting the supplied data in a more human readable form.
#
# Original author: Kit Peters <cpeters@ucmo.edu>
#
# Base URL $URL$
# $Id$
# $Rev$
# Last modified by $Author$
# Last modified $Date$

use strict;
use warnings;

use feature qw/switch/;
use autodie;
use charnames q/:full/;
use English qw/-no_match_vars/;
use Text::CSV_XS;
use Time::Piece;
use Config::Tiny;
use Readonly;
use Win32::OLE;
use Win32::OLE::Const 'Microsoft Internet Controls';
use Win32::OLE::Variant;
use Carp qw/carp croak/;
use File::Spec;
use Clone qw/clone/;
use Cwd;
use Encode;
use File::Copy;
use Perl6::Form;
use File::Temp;
# Testing

our $VERSION = 1.3;

Readonly my $EMPTY                  => q{};
Readonly my $FUNCTION_NAME_POSITION => 3;
Readonly my $DEBUG                  => 1;
Readonly my $HTML_PREAMBLE =>
q{<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">};

# constants used with Internet Explorer OLE from Vc7/PlatformSDK/Include/MsHtmHst.h :
Readonly my $PRINT_DONTBOTHERUSER    => 1;
Readonly my $PRINT_WAITFORCOMPLETION => 2;

Readonly my %DEFAULTS => (
    '_' => {
        'channels_file'       => q/channels.ini/,
        'transmitter_log_dir' => undef,
        'printer_path'        => undef,
        'field_order'         => undef,
        'print_with_ie'       => 1,
        'header_file'         => 'header.txt',
        'footer_file'         => 'footer.txt',
        'log_file'            => 'log.txt',
    },
);

# This is an array for purposes of Windows compatibility
Readonly my @CONFIG_PATH => ( File::Spec->splitpath($PROGRAM_NAME) )[ 0 .. 1 ];

Readonly my $CONFIG_FILE =>
  File::Spec->rel2abs( q/moseley-log-beautifier.ini/, @CONFIG_PATH );

Readonly my $CONFIG => eval { get_configuration($CONFIG_FILE); } or do {
    my $message = qq/Error reading config: $EVAL_ERROR/;
    _log_write($message);
    croak($message);
};

Readonly my $CHANNELS => eval {
    get_channels(
        File::Spec->rel2abs( $CONFIG->{'_'}{'channels_file'}, @CONFIG_PATH ) );
} or do {
    my $message = qq/Error reading channels: $EVAL_ERROR/;
    _log_write($message);
    croak($message);
};

main();

exit 0;

sub main {
    my $transmitter_log = $ARGV[0]
      || File::Spec->catfile( $CONFIG->{'_'}{'transmitter_log_dir'},
        q/Log.txt/ );
    _log_write(qq/Begin run on file "$transmitter_log"/);

    my $fh;
    eval {
        open $fh, '<', $transmitter_log;
        1;
    } or _error_exit(qq/Failed to open TX log: $EVAL_ERROR/);

    my $processed_records = eval { _process_transmitter_log($fh); }
      or _error_exit(qq/Failed to process TX logs: $EVAL_ERROR/);
    eval {
        close $fh;
        1;
    } or _error_exit(qq/Failed to close TX log: $EVAL_ERROR/);

    my $log_date     = [ sort keys %{$processed_records} ]->[0];
    my $tabular_data = format_tabular($processed_records);

    my $record_count = eval {
        print_processed_logs(
            { 'log_date' => $log_date, 'data' => $processed_records } );
    };
    if ( !defined $record_count ) {
        _error_exit(qq/Failed to print TX logs: $EVAL_ERROR/);
    }
    elsif ( $record_count == 0 ) {
        _error_exit(
qq/Zero horizontal records created from raw TX log file "$transmitter_log"/
        );
    }
    else {
        _log_write(
qq/TX logs processed successfully for $log_date.  $record_count records./
        );
    }
    return 1;
}

# my $record_count = print_processed_logs({ 'log_date' => $date, 'log_data' => \@data });
#
# Print out processed logs.
#
# Expects args in a hashref with keys:
## "log_date" Date that will be printed on the first line of the printout, under the header
## "data" An arrayref of arrayrefs (i.e. 2-D arrayref)  This is the data that will be printed.
#
# Returns number of rows log data have been processed into
sub print_processed_logs {
    my $args = shift;

    if ( ref $args ne q/HASH/ ) {
        croak(
            sprintf q/Usage: %s <hashref>/,
            ( caller 0 )[$FUNCTION_NAME_POSITION]
        );
    }
    foreach my $required_key (qw/log_date data/) {
        if ( !$args->{$required_key} ) {
            croak(qq/Missing required key '$required_key' in args/);
        }
    }

    if ( $CONFIG->{'_'}{'print_with_ie'} ) {
        _print_with_internet_explorer($args);
        _debug(q/Printed with IE/);
    }
    else {
        _print_as_text($args);
        _debug(q/Printed as text/);
    }
    _debug( q/Returning record count of / . scalar @{ $args->{'data'} } );
    return scalar @{ $args->{'data'} };
}

# my $tabular = format_tabular(\%records)
sub format_tabular {
    my $horizontal_records = shift;

    if ( ref $horizontal_records ne q/HASH/ ) {
        croak(
            sprintf q/Usage: %s <hashref>/,
            ( caller 0 )[$FUNCTION_NAME_POSITION]
        );
    }

    my @output_fields =
      map { $_->{'Description'} }
      @{$CHANNELS}{ @{ $CONFIG->{'_'}{'field_order'} } };

    my $output_formats = {};
    foreach my $key ( keys %{$CHANNELS} ) {
        my $field_name = $CHANNELS->{$key}{'Description'};
        my $units      = $CHANNELS->{$key}{'Units'};
        $output_formats->{$field_name} = $units;
    }

    my @tabular =
      ( [ q|Time|, @output_fields ] );    # initialize w/ column headings
    foreach my $timestamp ( sort keys %{$horizontal_records} ) {
        my $horizontal_record = $horizontal_records->{$timestamp};

        # Extract the time portion of the timestamp
        my ($time) = $timestamp =~ m/(\d{2}:\d{2}:\d{2})/xsm;

        # Add units to the tabular data
        foreach my $field_name (@output_fields) {
            my $unit = $output_formats->{$field_name};

            given ($unit) {
                when (/^none/ixsm) {
                    next;
                }
                when (/^bool/ixsm) {
                    $horizontal_record->{$field_name} =
                      ( $horizontal_record->{$field_name} ) ? 'YES' : 'NO';
                }
                when (/^percent/ixsm) {
                    $horizontal_record->{$field_name} .= q{%};
                }
                when (/^deg/ixsm) {
                    $horizontal_record->{$field_name} .= q{°};
                }
                default {
                    $horizontal_record->{$field_name} .= uc substr $unit, 0, 1;
                }
            }
        }
        push @tabular, [ $time, @{$horizontal_record}{@output_fields} ];
    }

    return \@tabular;
}

# Expects arguments as a hashref with the keys:
# # log_date: Date of the log
# # data: an arrayref of arrayrefs.  First line is treated as column headings, following lines are treated as data.
#
# A double horizontal rule will be added between the column headings and the data.
#
sub _print_with_internet_explorer {
    my $args = shift;

    if ( ref $args ne q/HASH/ ) {
        croak(
            sprintf q/Usage: %s <hashref>/,
            ( caller 0 )[$FUNCTION_NAME_POSITION]
        );
    }
    foreach my $required_key (qw/log_date data/) {
        if ( !$args->{$required_key} ) {
            croak(qq/Missing required key '$required_key' in args/);
        }
    }
    my $html     = _generate_html($args);
    my $temp     = File::Temp->new( q{SUFFIX} => q{.html} );
    my $filename = $temp->filename;

    binmode $temp, q{:crlf};
    $temp->print($html) or croak(qq/Failed to write to tempfile: $OS_ERROR/);
    $temp->flush;

    my $IE = Win32::OLE->new('InternetExplorer.Application')
      or croak( q/Failed to instantiate Internet Explorer for printing: /
          . Win32::OLE->LastError );
    $IE->Navigate($filename);

    while ( $IE->{q/Busy/} ) {
        sleep 1;
    }

    # Prints the active document in IE
    $IE->ExecWB( OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER,
        Variant( VT_I2, $PRINT_WAITFORCOMPLETION | $PRINT_DONTBOTHERUSER ) );

    $IE->Quit();
    $temp->close;

    return 1;
}

sub _generate_html {
    my $args = shift;

    my $header = _slurp_file( $CONFIG->{'_'}{'header_file'} );
    my $footer = _slurp_file( $CONFIG->{'_'}{'footer_file'} );

    $header =~ s/\n/<br \/>\n/gxsm;
    $footer =~ s/\n/<br \/>\n/gxsm;

    my @column_headings = @{ shift $args->{'data'} };
    my @rows            = @{ $args->{'data'} };

    my $html = <<"END";
$HTML_PREAMBLE
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title></title>
        <style type="text/css">
            td {
                text-align: right;
            }
            .header {
                text-align: center;
            }
        </style>
    </head>
    <body>
        <p class="header">$header</p>
        <table>           
END
    $html .= q{<tr>};
    $html .= join $EMPTY, map { qq{<th>$_</th>} } @column_headings;
    $html .= qq{</tr>\n};

    foreach my $row (@rows) {
        $html .= q{<tr>};
        $html .= join $EMPTY, map { qq{<td>$_</td>} } @{$row};
        $html .= qq{</tr>\n};
    }
    $html .= <<"END";
        </table>
        <p>$footer</p>
    </body>
</html> 
END
    return $html;
}

sub _print_as_text {
    my $args = shift;

    if ( ref $args ne q/HASH/ ) {
        croak(
            sprintf q/Usage: %s <hashref>/,
            ( caller 0 )[$FUNCTION_NAME_POSITION]
        );
    }
    foreach my $required_key (qw/log_date data/) {
        if ( !$args->{$required_key} ) {
            croak(qq/Missing required key '$required_key' in args/);
        }
    }

    my $header = _slurp_file( $CONFIG->{'_'}{'header_file'} );
    my $footer = _slurp_file( $CONFIG->{'_'}{'footer_file'} );

    my @column_headings = @{ shift $args->{'data'} };
    my @rows            = @{ $args->{'data'} };

    my $header_format = join q{|}, (q/{]]]][[[[}/) x scalar @column_headings;
    my $field_format  = join q{|}, (q/{]]]]]]]]}/) x scalar @column_headings;

    # formatting starts with headers followed by double line
    my @format_data = ( $header_format, @column_headings, );
    push @format_data, join q{|}, (q/==========/) x scalar @column_headings;
    foreach my $row (@rows) {
        push @format_data, ( $field_format, @{$row} );
    }
    my $text = form @format_data;

    my ( $fh, $tempfile ) = File::Temp::tempfile;
    $fh->binmode(q{:crlf});
    $fh->print($text) or croak(qq/Failed to write to tempfile: $OS_ERROR/);
    close $fh;

    File::Copy::copy( $tempfile, $CONFIG->{'_'}{'printer_path'} )
      or croak(qq/Failed to print file: $OS_ERROR/);
    unlink $tempfile;

    return 1;
}

sub _slurp_file {
    my $file = shift;
    open my $fh, '<', $file;
    my $text = do { local $INPUT_RECORD_SEPARATOR = undef; <$fh> };
    close $fh;

    return $text;
}

sub _process_transmitter_log {
    my $fh = shift;

    my $csv = Text::CSV_XS->new( { q/allow_whitespace/ => 1, q/binary/ => 1 } );

    my @column_names       = @{ $csv->getline($fh) };
    my $horizontal_records = {};
    my $vertical_record    = {};
    $csv->bind_columns( \@{$vertical_record}{@column_names} );
    while ( my $result = $csv->getline($fh) ) {
        if ( !defined $result && !$csv->eof ) {
            my ( $code, $message, $position, $record_num ) = $csv->error_diag();
            croak(
qq/Failed to process TX log: $message at record $record_num, character $position/
            );
        }
        my $time = $vertical_record->{'Time'};
        my $date = $vertical_record->{'Date'} . q{/} . localtime->year;
        my $value =
          $vertical_record->{'Current Value'} + 0;   # coerce this into a number
        my $key =
            $vertical_record->{'Type of Signal'}
          . $vertical_record->{'Channel number'};
        my $field_name = $CHANNELS->{$key}{'Description'}
          || qq/Channel $vertical_record->{'Channel number'}/;

        my $timestamp = qq/$date $time/;
        $horizontal_records->{$timestamp}{$field_name} = $value;
    }

    return $horizontal_records;
}

sub _log_write {
    my $message = shift;
    open my $fh, q{>>}, $CONFIG->{'_'}{'log_file'};
    my $timestamp = localtime->strftime('%c');
    my $ret       = $fh->print(qq/[$timestamp] $message\n/);
    if ( !$ret ) {
        croak(q/Failed to print to log/);
    }
    close $fh;
    return 1;
}

sub _debug {
    my $message = shift;
    if ($DEBUG) {
        _log_write(qq/DEBUG: $message/);
    }
    return 1;
}

sub _error_exit {
    my $message = shift;
    _log_write($message);
    croak($message);
}

sub get_channels {
    my $channels_file         = shift;
    my $channels_config_final = {};
    my $channels_config       = Config::Tiny->read($channels_file);
    if ( !$channels_config ) {
        croak( qq/Failed to read channels config "$channels_file": /
              . Config::Tiny->errstr );
    }
    foreach my $channel ( keys %{$channels_config} ) {
        if ( $channel =~ /^channel/ixsm ) {
            my ($channel_number) = $channel =~ m/channel\s+([[:alnum:]]+)/ixsm;
            $channels_config_final->{ uc $channel_number } =
              $channels_config->{$channel};
        }
        else {
            croak(
qq/Malformed key "[$channel]" in channels config "$channels_file"/
            );
        }
    }
    return $channels_config_final;
}

sub get_configuration {
    my $config_file = shift;
    my $config      = Config::Tiny->read($CONFIG_FILE);
    if ( !$config ) {
        croak( qq/Failed to read configuration file $config_file: /
              . Config::Tiny->errstr );
    }
    if ( $config->{'_'}{'field_order'} ) {
        my @field_order = split /\s+/xsm, $config->{'_'}{'field_order'};
        $config->{'_'}{'field_order'} = \@field_order;
    }

    foreach my $key ( keys %DEFAULTS ) {
        foreach my $subkey ( keys %{ $DEFAULTS{$key} } ) {
            if (   !defined $DEFAULTS{$key}{$subkey}
                && !$config->{$key}{$subkey} )
            {
                croak(qq/Required key "$key" not present in config file/);
            }
            if ( !defined $config->{$key}{$subkey}
                || $config->{$key}{$subkey} eq q{} )
            {
                $config->{$key}{$subkey} = clone( $DEFAULTS{$key}{$subkey} );
            }
        }
    }
    return $config;
}
__END__

=pod

=head1 NAME

moseley-log-beautifier.pl - A script to take the output from Moseley's CommServer, reformat it into a 
more readable form, and print the reformatted output. 

=head1 SYNOPSIS

Run without arguments, it will process "Log.txt" in the transmitter logs directory.

=head1 AUTHOR

Kit Peters <cpeters@ucmo.edu>

=head1 ACKNOWLEDGEMENTS

My thanks go to all the friendly folks at Stack Overflow, particularly "ikegami" and "daotoad", who so
patiently answered all my weird questions.

=head1 BUGS AND LIMITATIONS

The only time specifying the printer will work is if you're printing as text.  Printing with IE always
goes to the default printer.

=head1 USAGE

To process the file "Log.txt" in the TX logs directory (specified in config file) 

perl moseley-log-beautifier.pl 

Run with a single argument, it will process the file specified on the command line.

perl moseley-log-beautifier.pl //path/to/Log02132013.txt

Run with multiple arguments, it will treat the first argument as the log file to process and 
ignore the rest of the arguments.  Don't do this; it's silly.

=head1 REQUIRED ARGUMENTS 

None

=head1 DESCRIPTION

This script is part of a larger system designed to automate KMOS's TV transmitter meter readings.  It 
depends upon log files generated by Moseley CommServer, which is the program that actually logs the meter 
readings.  Configuration of CommServer is beyond the scope of this document.

=head1 CONFIGURATION

The program is configured by a configuration file "moseley-log-beautifier.ini", which is expected to be located in
the same directory as the program is run from.  Channel configuration is stored in a file, "channels.ini" 
(note that this can be changed in the config file), also expected (by default) to be in the same directory
as the program is run from.

=head2 CONFIG FILE OPTIONS

=over 4

=item channels_file

This specifies the location of the channel definitions file.  It is expected to be in .ini format, and 
each entry is expected to be of the form

    [Channel XNN]
        Description=Some description here...
        Units=[Degrees|Watts|Amps|Volts|Percentage|None]
    
Where B<X> is one of "T" (for telemetry channels) or "S" (for status channels), and B<NN> is the channel
number.  Description is a free-form string.  Units will be printed out with the proper abbreviation, e.g. 
"W" for watts, "%" for percentage, and E<deg> for degrees.

=item transmitter_log_dir

This specifies the location from which log files from CommServer will be read.

=item field_order

Order of fields to be displayed in the output file.  These fields should be in the same format as the 
channel definitions, e.g. 'T33' for telemetry channel 33.

Example: "field_order=T33 T34 T41 T48 T32 S1" will display, in order, telemetry channels 33, 34, 41, 48, 
and 32, and status channel 1.

=item printer_path

Path to the output printer. This is currently ignored if print_with_ie is set.

=item print_with_ie

If this is set, output will be generated in HTML and printed with Internet Explorer to the host computer's
default printer.

=item header_file

The contents of this file will be printed before the table of meter readings is printed.

=item footer_file

The contents of this file will be printed after the table of meter readings is printed.

=back

=head1 OPTIONS

Specify a specific file to process by specifiying the full path to the file on the command line

=head1 DIAGNOSTICS

Some output is written to a log file that can be specified in the configuration.  In future revisions Windows 
native logging support may be added.

=head1 DEPENDENCIES

This program depends on log files generated by Moseley CommServer.

=head1 EXIT STATUS

0 on success, nonzero on failure.

=head1 INCOMPATIBILITIES

At present, this program is designed to run more or less on Windows.  It has not been tested on Linux 
or other Unices.

=head1 LICENSE AND COPYRIGHT

This program is copyright (c) 2013 by the University of Central Missouri.  Licensed under the same terms 
as Perl itself.

=cut
