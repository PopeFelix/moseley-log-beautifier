#!/usr/bin/perl
#
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
    my $logfile = File::Spec->catfile($CONFIG->{'_'}{'transmitter_log_dir'}, q/Log.txt/);

    open my $fh, '<', $logfile;
    my $processed_records = _process_transmitter_log($fh);
    close $fh;
   
    print_processed_logs($processed_records);
    return 1;
}

# TODO: make this create a Word doc or text document and print it out
sub print_processed_logs {
    my $processed_records = shift;

    my $tabular_data = _format_tabular($processed_records);

    if ($CONFIG->{'_'}{'print_with_word'}) {
        _print_with_word($tabular_data);
    } 
    else {
        _print_as_text($tabular_data);
    }
    return 1;
}

sub _format_tabular {
    my $horizontal_records = shift;
    
    if (ref $horizontal_records ne q/HASH/) {
        croak(sprintf(q/Usage: %s <hashref>/, (caller(0))[3]));
    }
    
    my @output_fields = map { $_->{'Description'} } @{$CHANNELS}{ @{ $CONFIG->{'_'}{'field_order'} } };

    my @tabular = ([q|Time|, @output_fields]); # initialize w/ column headings
    foreach my $timestamp (sort keys %{$horizontal_records}) {
        my $record = $horizontal_records->{$timestamp};
        push @tabular, [$timestamp, @{$record}{@output_fields}];
    }

    return \@tabular;
}

# Expects an arrayref of arrayrefs.  First line is treated as column headings, following lines are treated as data.  
# A single horizontal rule will be added between the column headings and the data.
sub _print_with_word {
    my $print_data = shift;

    if (ref $print_data ne q/ARRAY/) {
        croak(sprintf(q/Usage: %s <arrayref>/, (caller(0))[3]));
    }

    my $header = _read_header($CONFIG->{'_'}{'header_file'});
    my $csv = Text::CSV_XS->new({ sep_char => ',' });

    my @column_headings = @{shift $print_data};
    my @rows = @{$print_data};

    my $word = Win32::OLE->new('Word.Application', 'Quit');
    my $doc = $word->Documents->Add({'Visible' => 1});
    my $select = $word->Selection;
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
    @{$table->Rows->First->Borders(wdBorderBottom)}{qw/LineStyle LineWidth/} = (wdLineStyleDouble, wdLineWidth100pt);
    $doc->SaveAs({ 'Filename' => Cwd::getcwd . '/test.doc' });
#    $doc->PrintOut();
    $doc->Close({ 'SaveChanges' => wdDoNotSaveChanges });
    $word->Quit();
    return 1;
}

sub _read_header {
    my $header_file = shift;

    open my $fh, '<', $header_file;
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

        $horizontal_records->{$time}{$field_name} = $value;
    }
    return $horizontal_records;
}

sub get_channels {
    my $channels_file = shift;
    $DB::single = 1;   
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
    $DB::single = 1;
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
