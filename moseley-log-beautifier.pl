#!/usr/bin/perl -- ## no critic (RequireRcsKeywords)
# Explanation: Source control for this project is via Git.

# A script to beautify automatic log output from Moseley CommServer by
# presenting the supplied data in a more human readable form.
#
# Original author: Kit Peters <cpeters@ucmo.edu>
#

use strict;
use warnings;
use Carp qw/carp croak/;
use Readonly;
use English qw/-no_match_vars/;
use Config::Tiny;
use Readonly;
use Win32::OLE;
use Win32::OLE::Const 'Microsoft Internet Controls';
use Win32::OLE::Variant;
use File::Spec;
use Clone qw/clone/;
use Cwd;
use Sys::Syslog qw/:standard :macros/;
use File::Temp;
use Moseley::LogBeautifier;
use Getopt::Long qw/:config auto_help/;
use Pod::Usage;

our $VERSION = 2.1;

Readonly my $EMPTY          => q{};
Readonly my $LOG_FACILITY   => Sys::Syslog::LOG_USER;
Readonly my $LOG_IDENTIFIER => q/moseley-log-beautifier/;
Readonly my $LOG_OPTIONS    => $EMPTY;

BEGIN {    # Start logging immediately
    openlog( $LOG_IDENTIFIER, $LOG_OPTIONS, $LOG_FACILITY );
}

Readonly my $FUNCTION_NAME_POSITION => 3;

# constants used with Internet Explorer OLE from Vc7/PlatformSDK/Include/MsHtmHst.h :
Readonly my $PRINT_DONTBOTHERUSER    => 1;
Readonly my $PRINT_WAITFORCOMPLETION => 2;

Readonly my %DEFAULTS => (
    'channels_file'       => q/channels.ini/,
    'transmitter_log_dir' => undef,
    'printer_path'        => undef,
    'field_order'         => undef,
    'print_with_ie'       => 1,
    'header_file'         => 'header.txt',
    'footer_file'         => 'footer.txt',
);

# This is an array for purposes of Windows compatibility
Readonly my @CONFIG_PATH => ( File::Spec->splitpath($PROGRAM_NAME) )[ 0 .. 1 ];

Readonly my $CONFIG_FILE =>
  File::Spec->rel2abs( q/moseley-log-beautifier.ini/, @CONFIG_PATH );

Readonly my $CONFIG => eval { get_configuration($CONFIG_FILE); } or do {
    my $message = qq/Error reading config: $EVAL_ERROR/;
    _error_exit($message);
};
$DB::single = 1;

eval {
    GetOptions(
        $CONFIG,                 'channels_file=s',
        'transmitter_log_dir=s', 'printer_path=s',
        'field_order=s',         'print_with_ie=s',
        'header_file=s',         'footer_file=s',
        'transmitter_log=s',
    );
};
if ($EVAL_ERROR) {
    pod2usage($EVAL_ERROR);
}


if ( $ARGV[0] ) {
    $CONFIG->{'transmitter_log'} = $ARGV[0];
}

main();

exit 0;

# Main program thread
sub main {

    my $transmitter_log = $CONFIG->{'transmitter_log'}
      || File::Spec->catfile( $CONFIG->{'transmitter_log_dir'}, q/Log.txt/ );
    syslog( LOG_INFO, qq/Begin run on file "$transmitter_log"/ );

    my $beautifier = eval {
        Moseley::LogBeautifier->new(
            {
                'filename'      => $transmitter_log,
                'channels_file' => $CONFIG->{'channels_file'},
                'header_file'   => $CONFIG->{'header_file'},
                'footer_file'   => $CONFIG->{'footer_file'},
                'field_order'   => $CONFIG->{'field_order'},
            }
        );
    } or error_exit(qq/Failed to instantiate LogBeautifier: $EVAL_ERROR/);

    my $record_count;

    if ( $CONFIG->{'print_with_ie'} ) {
        my $html = $beautifier->generate_html_output();
        $record_count = print_with_internet_explorer($html);
        syslog( LOG_DEBUG, qq/Printed $record_count records with IE/ );
    }
    else {
        my $text = $beautifier->generate_text_output();
        $record_count = print_as_text($text);
        syslog( LOG_DEBUG, qq/Printed $record_count records as text/ );
    }
    return defined $record_count;
}

# print the given HTML with Internet explorer
#
# Takes a single argument, which is the HTML text to print.
# die()'s on error
sub print_with_internet_explorer {
    my $html = shift or croak(q/Usage: print_with_internet_explorer(<html>)/);

    my $temp = File::Temp->new( q{SUFFIX} => q{.html} );
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

# Write the error message to the syslog and exit
#
# Takes a single argument: message to write to the syslog
sub error_exit {
    my $message = shift;
    syslog( LOG_ERR, $message );
    croak($message);
}

# Read configuration from the specified config file
# 
# Takes a single argument, which is the location of the config file to read.
# Returns the parsed configuration
# die()'s on error
sub get_configuration {
    my $config_file = shift;

    my $config = Config::Tiny->read($config_file)->{'_'};
    if ( !$config ) {
        croak( qq/Failed to read configuration file $config_file: /
              . Config::Tiny->errstr );
    }
    if ( $config->{'field_order'} ) {
        my @field_order = split /\s+/xsm, $config->{'field_order'};
        $config->{'field_order'} = \@field_order;
    }

    # config file trumps defaults
    foreach my $key ( keys %DEFAULTS ) {
        if ( !defined $DEFAULTS{$key} && !$config->{$key} ) {
            croak(qq/Required key "$key" not present in config file/);
        }
        if ( !defined $config->{$key} || $config->{$key} eq $EMPTY ) {
            $config->{$key} = clone( $DEFAULTS{$key} );
        }
    }

    return $config;
}

END {
    closelog;
}
__END__

=pod

=head1 NAME

moseley-log-beautifier.pl - A script to print readable log files from a Moseley remote control system

=head1 SYNOPSIS

moseley-log-beautifier.pl [options] [log file]

  Options:
    --channels_file         location of the channels file
    --transmitter_log_dir   directory where TX logs are stored
    --field_order           order to print data fields in 
    --print_with_ie         use Internet Explorer to print (disables --printer_path)
    --printer_path          printer to use (only when --print-with-ie is not used)
    --header_file           text file containing header to be printed with each printout
    --footer_file           text file containing footer to be printed with each printout
    --transmitter_log       Moseley log file to process      

  All options override the corresponding entries in the config file

  Run without arguments, it will process "Log.txt" in the transmitter logs directory.

=head1 OPTIONS

=over 8

=item B<--channels_file>

This specifies the location of the channel definitions file.  It is expected to be in .ini format.  See L<"CHANNELS FILE"> below

=item B<--transmitter_log_dir>

This specifies the location from which log files from CommServer will be read.

=item B<--field_order>

Order of fields to be displayed in the output file.  These fields should be in the same format as the 
channel definitions, e.g. 'T33' for telemetry channel 33.

Example: "field_order=T33 T34 T41 T48 T32 S1" will display, in order, telemetry channels 33, 34, 41, 48, 
and 32, and status channel 1.

=item B<--printer_path>

Path to the output printer. This is currently ignored if print_with_ie is set.

=item B<--print_with_ie>

If this is set, output will be generated in HTML and printed with Internet Explorer to the host computer's
default printer.

=item B<--header_file>

The contents of this file will be printed before the table of meter readings is printed.

=item B<--footer_file>

The contents of this file will be printed after the table of meter readings is printed.

=back

=head1 CONFIGURATION

The program is configured by a configuration file "moseley-log-beautifier.ini", which is expected to be located in the same directory as the program is run from.  Channel configuration is stored in a file, "channels.ini" (note that this can be changed in the config file), also expected (by default) to be in the same directory as the program is run from.  Note that any options specified on the command line will override the corresponding value in the config file.

=head2 CONFIG FILE OPTIONS

=over 8

=item channels_file

This specifies the location of the channel definitions file. See L<"CHANNELS FILE"> below.

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

=head2 CHANNELS FILE

The channels file is expected to be in .ini format, and each entry is expected to be of the form

    [Channel XNN]
        Description=Some description here...
        Units=[Degrees|Watts|Amps|Volts|Percentage|None]
    
Where B<X> is one of "T" (for telemetry channels) or "S" (for status channels), and B<NN> is the channel
number.  Description is a free-form string.  Units will be printed out with the proper abbreviation, e.g. 
"W" for watts, "%" for percentage, and E<deg> for degrees.

=head1 AUTHOR

Kit Peters <cpeters@ucmo.edu>

=head1 ACKNOWLEDGEMENTS

My thanks go to all the friendly folks at Stack Overflow, particularly "ikegami" and "daotoad", who so
patiently answered all my weird questions.

=head1 BUGS AND LIMITATIONS

The only time specifying the printer will work is if you're printing as text.  Printing with IE always
goes to the default printer.

Under Windows, log messages will show up in the Event Log with a warning message such as "The description 
for Event ID 157 from source moseley-log-beautifier.pl [SSW:1.0.1] cannot be found. Either the component 
that raises this event is not installed on your local computer or the installation is corrupted. You can
install or repair the component on the local computer."  
I believe this to be a bug in Sys::Syslog, and I have reported it as such on CPAN.

=head1 REQUIRED ARGUMENTS 

None

=head1 DESCRIPTION

This script is part of a larger system designed to automate KMOS's TV transmitter meter readings.  It 
depends upon log files generated by Moseley CommServer, which is the program that actually logs the meter 
readings.  Configuration of CommServer is beyond the scope of this document.

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
