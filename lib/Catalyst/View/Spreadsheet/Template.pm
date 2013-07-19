package Catalyst::View::Spreadsheet::Template;
use Moose;
use namespace::autoclean;
# ABSTRACT: render Spreadsheet::Template templates in Catalyst

use Path::Class::File;
use Try::Tiny;

use Spreadsheet::Template;

extends 'Catalyst::View';

=head1 SYNOPSIS

  package MyApp::View::Spreadsheet::Template;
  use Moose;

  extends 'Catalyst::View::Spreadsheet::Template';

=head1 DESCRIPTION

This module provides a L<Catalyst::View> for L<Spreadsheet::Template>.

=cut

=attr path

Template search path. Defaults to C<< [ $c->path_to('root') ] >>.

=cut

has path => (
    traits    => ['Array'],
    isa       => 'ArrayRef[Path::Class::Dir]',
    writer    => 'set_path',
    predicate => 'has_path',
    handles   => {
        path => 'elements',
    },
);

=attr processor_class

The C<processor_class> to pass through to the L<Spreadsheet::Template> object.

=cut

has processor_class => (
    is      => 'ro',
    isa     => 'Str',
    default => 'Spreadsheet::Template::Processor::Xslate',
);

=attr writer_class

The C<writer_class> to pass through to the L<Spreadsheet::Template> object.

=cut

has writer_class => (
    is      => 'ro',
    isa     => 'Str',
    default => 'Spreadsheet::Template::Writer::XLSX',
);

=attr template_extension

The extension to use for template files. Defaults to C<json>.

=cut

has template_extension => (
    is      => 'ro',
    isa     => 'Str',
    default => 'json',
);

=attr catalyst_var

The variable name to use for the Catalyst context object in the template.
Defaults to C<c>.

=cut

has catalyst_var => (
    is      => 'ro',
    isa     => 'Str',
    default => 'c',
);

has renderer => (
    is  => 'rw',
    isa => 'Spreadsheet::Template',
);

sub ACCEPT_CONTEXT {
    my $self = shift;
    my ($c) = @_;

    $self->renderer(
        Spreadsheet::Template->new(
            processor_class => $self->processor_class,
            writer_class    => $self->writer_class,
        )
    );

    $self->set_path([ $c->path_to('root') ]) unless $self->has_path;

    return $self;
}

my %content_types = (
    xlsx => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    xls  => 'application/vnd.ms-excel',
);

sub process {
    my $self = shift;
    my ($c) = @_;

    return try {
        my $content = $self->render($c);
        $c->response->content_type($content_types{$self->_extension});
        $c->response->body($content);
        1;
    }
    catch {
        my $error = "Couldn't render template: $_";
        $c->log->error($error);
        $c->error($error);
        0;
    };
}

sub render {
    my $self = shift;
    my ($c) = @_;

    $self->renderer->render(
        scalar($self->template_file($c)->slurp),
        {
            %{ $c->stash },
            $self->catalyst_var => $c,
        }
    );
}

sub template_file {
    my $self = shift;
    my ($c) = @_;

    my $file = $c->stash->{template}
            || $c->action . '.' . $self->template_extension;

    for my $dir ($self->path) {
        my $full_path = $dir->file($file);
        if (-e $full_path) {
            return $full_path;
        }
    }

    die "Couldn't find template file $file in " . join(", ", $self->path);
}

sub _extension {
    my $self = shift;

    (my $extension = lc($self->writer_class)) =~ s/.*:://;

    return $extension;
}

__PACKAGE__->meta->make_immutable;
no Moose;

=head1 BUGS

No known bugs.

Please report any bugs to GitHub Issues at
L<https://github.com/doy/catalyst-view-spreadsheet-template/issues>.

=head1 SEE ALSO

L<Spreadsheet::Template>

L<Catalyst::View::Excel::Template::Plus>

=head1 SUPPORT

You can find this documentation for this module with the perldoc command.

    perldoc Catalyst::View::Spreadsheet::Template

You can also look for information at:

=over 4

=item * MetaCPAN

L<https://metacpan.org/release/Catalyst-View-Spreadsheet-Template>

=item * RT: CPAN's request tracker

L<http://rt.cpan.org/NoAuth/Bugs.html?Dist=Catalyst-View-Spreadsheet-Template>

=item * Github

L<https://github.com/doy/catalyst-view-spreadsheet-template>

=item * CPAN Ratings

L<http://cpanratings.perl.org/d/Catalyst-View-Spreadsheet-Template>

=back

=head1 SPONSORS

Parts of this code were paid for by

=over 4

=item Socialflow L<http://socialflow.com>

=back

=begin Pod::Coverage

  render
  template_file

=end Pod::Coverage

=cut

1;
