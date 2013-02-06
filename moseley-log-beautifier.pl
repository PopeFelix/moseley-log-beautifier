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

use autodie;
use English qw/-no_match_vars/;
use Text::CSV_XS;
use Time::Piece;
use Config::Tiny;
use Readonly;
use Win32::OLE;
use Win32::OLE::Const 'Microsoft Word';
use Carp qw/carp croak/;
use File::Spec;
use Clone qw/clone/;
use Cwd;

Readonly my %DEFAULTS => (
    '_' => {
        'channels_file' => q/channels.ini/,
        'transmitter_log_dir' => q|Z:|,
        'printer_path' => q|//153.91.87.132/HP DeskJet 712C|,
        'field_order'=> [qw/T33 T34 T41 T48 T32 S1/],
        'print_with_word' => 1,
        'header_file' => 'header.txt',
        'footer_file' => 'footer.txt',
    },
);
Readonly my $CONFIG_FILE          => q/moseley-log-beautifier.ini/;
Readonly my $CONFIG               => get_configuration($CONFIG_FILE);
if (!$CONFIG) {
    die qq/Failed to load config file "$CONFIG_FILE": / . Config::Tiny->errstr;
}
Readonly my $CHANNELS             => get_channels($CONFIG->{'_'}{'channels_file'});

main();
1;

sub main {
    # FIXME: figure out if I'm opening the file dated the night before or what.
    my $logfile = $ARGV[0] || File::Spec->catfile($CONFIG->{'_'}{'transmitter_log_dir'}, q/Log.txt/);

    open my $fh, '<', $logfile;
    my $processed_records = _process_transmitter_log($fh);
    close $fh;
   
    my $log_date = [sort keys %{$processed_records}]->[0];
    print_processed_logs({ 'log_date' => $log_date, 'log_data' => $processed_records});
    return 1;
}

sub print_processed_logs {
    my $args = shift;
    
    if (ref $args ne q/HASH/) {
        croak(sprintf(q/Usage: %s <hashref>/, (caller(0))[3]));
    }
    foreach my $required_key (qw/log_date log_data/) {
        if (!$args->{$required_key}) {
            croak(qq/Missing required key '$required_key' in args/);
        }
    }

    my $tabular_data = _format_tabular($args->{'log_data'});

    if ($CONFIG->{'_'}{'print_with_word'}) {
        _print_with_word({'log_date' => $args->{'log_date'}, 'data' => $tabular_data});
    } 
    else {
        _print_as_text({'log_date' => $args->{'log_date'}, 'data' => $tabular_data});
    }
    return 1;
}

sub _format_tabular {
    my $horizontal_records = shift;
    
    if (ref $horizontal_records ne q/HASH/) {
        croak(sprintf(q/Usage: %s <hashref>/, (caller(0))[3]));
    }
    
    my @output_fields = map { $_->{'Description'} } @{$CHANNELS}{ @{ $CONFIG->{'_'}{'field_order'} } };
    my $output_formats = {};
    foreach my $key (keys %{$CHANNELS}) {
        my $field_name = $CHANNELS->{$key}{'Description'};
        my $units = $CHANNELS->{$key}{'Units'};
        $output_formats->{$field_name} = $units;
    }
    my @tabular = ([q|Time|, @output_fields]); # initialize w/ column headings
    foreach my $timestamp (sort keys %{$horizontal_records}) {
        my $record = $horizontal_records->{$timestamp};
      
        (undef, my $time) = split /\s/, $timestamp, 2; # throw away the date portion of the timestamp
        # Add units to the tabular data
        foreach my $field_name (@output_fields) {
            my $unit = $output_formats->{$field_name};
            if ($unit =~ /none/ixsm) {
                next;
            }
            elsif ($unit =~ /bool/ixsm) {
                $record->{$field_name} = ($record->{$field_name}) ? 'YES' : 'NO';
            } 
            elsif ($unit =~ /percent/ixsm)  {
                $record->{$field_name} .= '%';
            }
            elsif ($unit =~ /deg/ixsm) {
                $record->{$field_name} .= qq/\xB0/; # degree symbol - NOTE: UTF-8
            }
            else {
                $record->{$field_name} .= uc(substr($unit, 0, 1));
            }
        }
        push @tabular, [$time, @{$record}{@output_fields}];
    }

    return \@tabular;
}

# Expects arguments as a hashref with the keys:
# # log_date: Date of the log
# # data: an arrayref of arrayrefs.  First line is treated as column headings, following lines are treated as data.  
#
# A double horizontal rule will be added between the column headings and the data.
sub _print_with_word {
    my $args = shift;
    
    if (ref $args ne q/HASH/) {
        croak(sprintf(q/Usage: %s <hashref>/, (caller(0))[3]));
    }
    foreach my $required_key (qw/log_date data/) {
        if (!$args->{$required_key}) {
            croak(qq/Missing required key '$required_key' in args/);
        }
    }

    my $header = _slurp_file($CONFIG->{'_'}{'header_file'});
    my $footer = _slurp_file($CONFIG->{'_'}{'footer_file'});
    my $csv = Text::CSV_XS->new({ 'sep_char' => ',', 'binary' => 1, 'quote_char' => undef });

    my @column_headings = @{shift $args->{'data'}};
    my @rows = @{$args->{'data'}};

    my $word = Win32::OLE->new('Word.Application', 'Quit');
    my $doc = $word->Documents->Add();
    my $select = $word->Selection;

    $select->ParagraphFormat->{'SpaceAfter'} = 0;
    $select->TypeText({'Text' => qq/$header\n\n/,});
    $select->BoldRun();
    $select->ParagraphFormat->{'Alignment'} = wdAlignParagraphRight;
    $select->TypeText({'Text' => Time::Piece->strptime($args->{'log_date'}, q|%m/%d/%Y %H:%M:%S|)->strftime(qq/%A %B %d %Y\n\n/)});
    $select->BoldRun();

    $csv->combine(@column_headings);
    $select->InsertAfter($csv->string);
    $select->InsertParagraphAfter;
    for my $row (@rows) {
        $csv->combine(@{$row});
        $select->InsertAfter($csv->string);
        $select->InsertParagraphAfter;
    }

    my $table = $select->ConvertToTable({'Separator' => wdSeparateByCommas});
    $table->Rows->First->Range->Font->{'Bold'} = 1;
    $table->Rows->First->Range->ParagraphFormat->{'Alignment'} = wdAlignParagraphCenter;
    @{$table->Rows->First->Borders(wdBorderBottom)}{qw/LineStyle LineWidth/} = (wdLineStyleDouble, wdLineWidth100pt);
    $doc->Paragraphs->Last->Format->{'Alignment'} = wdAlignParagraphLeft;
    $doc->Paragraphs->Last->Format->{'SpaceAfter'} = 0;
    $doc->Paragraphs->Last->Range->InsertAfter({'Text' => qq/\n$footer/}); 
    $doc->SaveAs({ 'Filename' => Cwd::getcwd . '/test.doc' });
#    $doc->PrintOut();
    $doc->Close({ 'SaveChanges' => wdDoNotSaveChanges });
    $word->Quit();
    return 1;
}

sub _slurp_file {
    my $file = shift;

    open my $fh, '<', $file;
    my $text = do { local ($/); <$fh> };
    close $fh;
    
    return $text;
}

sub _process_transmitter_log {
    my $fh = shift;
   
    my $csv = Text::CSV_XS->new( { q/allow_whitespace/ => 1, } );
    
    my @column_names = @{$csv->getline($fh)};
    my $horizontal_records = {};
    my $vertical_record = {};
    $csv->bind_columns(\@{$vertical_record}{@column_names});
    while ($csv->getline($fh)) {
        my $time = $vertical_record->{'Time'};
        my $date = $vertical_record->{'Date'} . '/' . localtime->year;
        my $value = $vertical_record->{'Current Value'} + 0; # coerce this into a number
        my $key = $vertical_record->{'Type of Signal'} . $vertical_record->{'Channel number'};
        my $field_name = $CHANNELS->{$key}{'Description'} || qq/Channel $vertical_record->{'Channel number'}/;

        my $timestamp = qq/$date $time/;
        $horizontal_records->{$timestamp}{$field_name} = $value;
    }
    return $horizontal_records;
}

sub get_channels {
    my $channels_file = shift;
    my $channels_config_final = {};
    my $channels_config = Config::Tiny->read($channels_file);
    foreach my $channel (keys %$channels_config) {
        if ($channel =~ /^channel/ixsm) {
            (my $channel_number = $channel) =~ s/channel ([[:alnum:]]+)/$1/ism;
            $channels_config_final->{$channel_number} = $channels_config->{$channel};
        }
    }
    return $channels_config_final;
}


sub get_configuration {
    my $config_file = shift;
    my $config = Config::Tiny->read($CONFIG_FILE);
    if (!$config) {
        croak(qq/Failed to read configuration file $config_file: / . Config::Tiny->errstr);
    }
    if ($config->{'_'}{'field_order'}) {
        my @field_order = split /\s+/, $config->{'_'}{'field_order'};
        $config->{'_'}{'field_order'} = \@field_order;
    }

    foreach my $key (keys %DEFAULTS) {
        foreach my $subkey (keys %{$DEFAULTS{$key}}) {
            $config->{$key}{$subkey} ||= clone($DEFAULTS{$key}{$subkey});
        }
    }
    return $config;
}
1;
