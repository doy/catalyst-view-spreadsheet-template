package Catalyst::View::Spreadsheet::Template;
use Moose;
use namespace::autoclean;

use Path::Class::File;
use Try::Tiny;

use Spreadsheet::Template;

extends 'Catalyst::View';

has renderer => (
    is  => 'rw',
    isa => 'Spreadsheet::Template',
);

has path => (
    traits    => ['Array'],
    isa       => 'ArrayRef[Path::Class::Dir]',
    writer    => 'set_path',
    predicate => 'has_path',
    handles   => {
        path => 'elements',
    },
);

has processor_class => (
    is      => 'ro',
    isa     => 'Str',
    default => 'Spreadsheet::Template::Processor::Xslate',
);

has writer_class => (
    is      => 'ro',
    isa     => 'Str',
    default => 'Spreadsheet::Template::Writer::XLSX',
);

has template_extension => (
    is      => 'ro',
    isa     => 'Str',
    default => 'json',
);

has catalyst_var => (
    is      => 'ro',
    isa     => 'Str',
    default => 'c',
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

1;
