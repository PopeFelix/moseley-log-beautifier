#!/usr/bin/perl -- ## no critic (RequireRcsKeywords)
# Explanation: Source control for this project is via Git.

package Moseley::LogBeautifier;

use Readonly;
use Carp qw/carp croak/;
use Moose;
use namespace::autoclean;
use Moose::Util::TypeConstraints;
use MooseX::Params::Validate;
use Config::Tiny;
use autodie;
use Text::CSV_XS;
use English qw/-no_match_vars/;
use feature qw/switch/;
use Perl6::Form;
use Time::Piece;
use File::Slurp;

Readonly my $EMPTY            => q{};
Readonly my $FIELD_WIDTH_TEXT => 8;
Readonly my $DEGREE_CHAR      => q{°};

our $VERSION = 1.0;

has '_log_date' => (
    'is'       => 'rw',
    'isa'      => 'Str',
    'init_arg' => undef,
);

has '_formatted_records' => (
    'is'       => 'rw',
    'isa'      => 'ArrayRef',
    'init_arg' => undef,
);

has '_channels' => (
    'is'       => 'rw',
    'isa'      => 'HashRef',
    'init_arg' => undef,
);

has 'filename' => (
    'is'       => 'ro',
    'isa'      => 'Str',
    'required' => 1,
);

has 'channels_file' => (
    'is'       => 'ro',
    'isa'      => 'Str',
    'required' => 1,
);

has 'field_order' => (
    'is'       => 'ro',
    'isa'      => 'ArrayRef',
    'required' => 1,
);

has 'header_file' => (
    'is'  => 'ro',
    'isa' => 'Str',
);

has 'footer_file' => (
    'is'  => 'ro',
    'isa' => 'Str',
);

sub BUILD {
    my $self = shift;
    my $args = shift;

    my $channels = $self->_parse_channels_file( $args->{'channels_file'} );
    $self->_channels($channels);

    open my $fh, '<', $self->filename;
    my $horizontal_records = $self->_generate_horizontal_records($fh);
    close $fh;

    my $log_date = [ sort keys %{$horizontal_records} ]->[0];
    $log_date =~ s/(\d{2}.\d{2}.\d{4}).+/$1/xsm;
    $self->_log_date($log_date);

    my $tabular = $self->_format_tabular($horizontal_records);
    $self->_formatted_records($tabular);
    return 1;
}

sub _generate_horizontal_records {
    my $self = shift;
    my $fh   = shift;

    my $csv = Text::CSV_XS->new( { q/allow_whitespace/ => 1, q/binary/ => 1 } );

    my @column_names = @{ $csv->getline($fh) };

    my $working_date;
    my $horizontal_records = {};
    my $vertical_record    = {};
    $csv->bind_columns( \@{$vertical_record}{@column_names} );
    while ( my $result = $csv->getline($fh) ) {
        if ( !defined $vertical_record && !$csv->eof ) {
            my ( $code, $message, $position, $record_num ) = $csv->error_diag();
            croak(
qq/Failed to process TX log: $message at record $record_num, character $position/
            );
        }
        my ( $time, $date, $value, $field_name );
        if ( $vertical_record->{'Type of alarm'} eq 'P' ) {    # periodic log
            $time = $vertical_record->{'Time'};
            $date = $vertical_record->{'Date'} . q{/} . localtime->year;

            $working_date ||= $date;

            my $key =
                $vertical_record->{'Type of Signal'}
              . $vertical_record->{'Channel number'};

            $field_name = $self->_channels->{$key}{'Description'}
              || qq/Channel $vertical_record->{'Channel number'}/;

            $value = $vertical_record->{'Current Value'};

            # if the value is all "?", no reading was taken
            if ( $vertical_record->{'Current Value'} =~ /^[?]+$/xsm ) {
                $value = q{N/A};
            }
        }
        else {
            $time = $vertical_record->{'Local Time'};
            $date =
              $working_date; # I don't think $working_date will ever be empty...

            my $first_field = $self->field_order->[0];

            $field_name = $self->_channels->{$first_field}{'Description'};
            $value      = qq/ALARM: $vertical_record->{'Description'}/;
        }
        my $timestamp = qq/$date $time/;
        $horizontal_records->{$timestamp}{$field_name} = $value;
    }

    return $horizontal_records;
}

sub _format_tabular {
    my $self               = shift;
    my $horizontal_records = shift;

    my @output_fields =
      map { $_->{'Description'} }
      @{ $self->_channels }{ @{ $self->field_order } };

    my $output_formats = {};
    foreach my $key ( keys %{ $self->_channels } ) {
        my $field_name = $self->_channels->{$key}{'Description'};
        my $units      = $self->_channels->{$key}{'Units'};
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
            if ( !defined $horizontal_record->{$field_name} ) {
                $horizontal_record->{$field_name} = $EMPTY;
                next;
            }

            if (   $horizontal_record->{$field_name} ne q{N/A}
                && $horizontal_record->{$field_name} !~ /^ALARM/xsm )
            {
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
                        $horizontal_record->{$field_name} .= $DEGREE_CHAR;
                    }
                    default {
                        $horizontal_record->{$field_name} .= uc substr $unit, 0,
                          1;
                    }
                }
            }
        }
        push @tabular, [ $time, @{$horizontal_record}{@output_fields} ];
    }

    return \@tabular;
}

sub _parse_channels_file {
    my $self                  = shift;
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

sub generate_html_output {
    my $self = shift;

    my $header = $EMPTY;
    if ( $self->header_file ) {
        $header = File::Slurp::read_file( $self->header_file );
        $header =~ s/\n/<br \/>\n/gxsm;
    }

    my $footer = $EMPTY;
    if ( $self->footer_file ) {
        $footer = File::Slurp::read_file( $self->footer_file );
        $footer =~ s/\n/<br \/>\n/gxsm;
    }

    my @column_headings = @{ shift $self->_formatted_records };
    my @rows            = @{ $self->_formatted_records };
    my $log_date        = $self->_log_date;
    my $html            = <<"END";
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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
        <p class="log_date">$log_date</p>
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
    $html =~ s/$DEGREE_CHAR/&deg;/gxsm;
    return $html;
}

sub generate_text_output {
    my $self = shift;

    my $header = $EMPTY;
    if ( $self->header_file ) {
        $header = File::Slurp::read_file( $self->header_file );
        $header =~ s/\n/<br \/>\n/gxsm;
    }

    my $footer = $EMPTY;
    if ( $self->footer_file ) {
        $footer = File::Slurp::read_file( $self->footer_file );
        $footer =~ s/\n/<br \/>\n/gxsm;
    }

    my @column_headings = @{ shift $self->_formatted_records };
    my @rows            = @{ $self->_formatted_records };

    my $header_field_format = q/{/
      . q{]} x ( $FIELD_WIDTH_TEXT / 2 )
      . q{[} x ( $FIELD_WIDTH_TEXT / 2 ) . q/}/;
    my $individual_field_format = q/{/ . q{]} x $FIELD_WIDTH_TEXT . q/}/;

    my $header_format = join q{|},
      ($header_field_format) x scalar @column_headings;
    my $date_format =
      q{ } x ( $FIELD_WIDTH_TEXT * scalar @column_headings ) . q/{>>>>>>>>>>}/;
    my $field_format = join q{|},
      ($individual_field_format) x scalar @column_headings;

    # formatting starts with headers followed by double line
    my @format_data =
      ( $date_format, $self->_log_date, $header_format, @column_headings, );
    push @format_data, join q{|}, (q/==========/) x scalar @column_headings;

    foreach my $row (@rows) {
        push @format_data, ( $field_format, @{$row} );
    }
    my $text = $header;
    $text .= form @format_data;
    $text .= $footer;

    return $text;
}

__PACKAGE__->meta->make_immutable;
1;
