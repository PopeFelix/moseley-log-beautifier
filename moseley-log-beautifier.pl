#!/usr/bin/perl -- ## no critic (RequireRcsKeywords)

# A script to beautify automatic log output from Moseley CommServer by
# presenting the supplied data in a more human readable form.
#
# Original author: Kit Peters <cpeters@ucmo.edu>
#

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

our $VERSION = 1.21;

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
Readonly my $CONFIG_FILE => q/moseley-log-beautifier.ini/;
Readonly my $CONFIG => eval { get_configuration($CONFIG_FILE); } or do {
    my $message = qq/Error reading config: $EVAL_ERROR/;
    _log_write($message);
    croak($message);
};

Readonly my $CHANNELS =>
  eval { get_channels( $CONFIG->{'_'}{'channels_file'} ); } or do {
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
    my $record_count = eval {
        print_processed_logs(
            { 'log_date' => $log_date, 'log_data' => $processed_records } );
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

sub print_processed_logs {
    my $args = shift;

    if ( ref $args ne q/HASH/ ) {
        croak(
            sprintf q/Usage: %s <hashref>/,
            ( caller 0 )[$FUNCTION_NAME_POSITION]
        );
    }
    foreach my $required_key (qw/log_date log_data/) {
        if ( !$args->{$required_key} ) {
            croak(qq/Missing required key '$required_key' in args/);
        }
    }

    my $tabular_data = _format_tabular( $args->{'log_data'} );
    my ($log_date) = $args->{'log_date'} =~ m/(\d{2}\/\d{2}\/\d{4})/xsm;
    if ( $CONFIG->{'_'}{'print_with_ie'} ) {
        _print_with_internet_explorer(
            { 'log_date' => $log_date, 'data' => $tabular_data } );
        _debug(q/Printed with Word/);
    }
    else {
        _print_as_text(
            { 'log_date' => $args->{'log_date'}, 'data' => $tabular_data } );
        _debug(q/Printed as text/);
    }
    _debug( q/Returning record count of / . scalar @{$tabular_data} );
    return scalar @{$tabular_data};
}

sub _format_tabular {
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
            .log_date {
            	text-align: right;
            }
        </style>
    </head>
    <body>
        <p class="header">$header</p>
        <p class="log_date">$args->{'log_date'}</p>
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
            if ( !defined $config->{$key}{$subkey}
                || $config->{$key}{$subkey} eq q{} )
            {
                $config->{$key}{$subkey} = clone( $DEFAULTS{$key}{$subkey} );
            }
        }
    }
    return $config;
}
