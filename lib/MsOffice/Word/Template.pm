package MsOffice::Word::Template;
use Moose;
use MooseX::StrictConstructor;
use Carp                           qw(croak);
use Template;

use namespace::clean -except => 'meta';


our $VERSION = '1.0';

has 'surgeon'         => (is => 'ro',   isa => 'MsOffice::Word::Surgeon', required => 1);
has 'data_color'      => (is => 'ro',   isa => 'Str',                     default  => "yellow");
has 'directive_color' => (is => 'ro',   isa => 'Str',                     default  => "green");
has 'template_config' => (is => 'ro',   isa => 'HashRef',                 default  => sub { {} });

has 'template_text'   => (is => 'bare', isa => 'Str',                     init_arg => undef);


#======================================================================
# BUILDING THE TEMPLATE
#======================================================================


# syntactic sugar for supporting ->new($surgeon) instead of ->new(surgeon => $surgeon)
around BUILDARGS => sub {
  my $orig  = shift;
  my $class = shift;

  if ( @_ == 1 && !ref $_[0] ) {
    return $class->$orig(surgeon => $_[0]);
  }
  else {
    return $class->$orig(@_);
  }
};


sub BUILD {
  my ($self) = @_;

  $self->{template_text} = $self->build_template_text;
}


sub template_fragment_for_run { # given a run node, build a template fragment
  my ($self, $run) = @_;

  my $props           = $run->props;
  my $data_color      = $self->data_color;
  my $directive_color = $self->directive_color;

  # if this run is highlighted in yellow or green, it must be translated into a TT2 directive
  # NOTE:  the translation code has much in common with Surgeon::Run::as_xml() -- maybe
  # part of the code could be shared in a future version
  if ($props =~ s{<w:highlight w:val="($data_color|$directive_color)"/>}{}) {
    my $color       = $1;
    my $xml         = $run->xml_before;

    my $inner_texts = $run->inner_texts;
    if (@$inner_texts) {
      $xml .= "<w:r>";
      $xml .= "<w:rPr>" . $props . "</w:rPr>" if $props;
      $xml .= "<w:t>[% ";
      $xml .= $_->literal_text . "\n" foreach @$inner_texts;
        # NOTE : adding "\n" because end of lines are used by templating modules
      $xml .= " %]";
      $xml .= "<!--TT2_directive-->" if $color eq $directive_color; # XML comment for marking TT2 directives -- used in regexes below
      $xml .= "</w:t>";
      $xml .= "</w:r>";
    }

    return $xml;
  }

  # otherwise this run is just regular MsWord content
  else {
    return $run->as_xml;
  }
}


sub build_template_text {
  my ($self) = @_;

  # regex for matching paragraphs that contain TT2 directives
  my $regex_paragraph = qr{
    <w:p               [^>]*>                  # start paragraph node
      <w:r             [^>]*>                  # start run node
        <w:t           [^>]*>                  # start text node
          (\[% .*? %\])   (*SKIP)              # template directive
          <!--TT2_directive-->                 # followed by an XML comment
        </w:t>                                 # close text node
      </w:r>                                   # close run node
    </w:p>                                     # close paragraph node
   }sx;

  # regex for matching table rows that contain TT2 directives
  my $regex_row = qr{
    <w:tr              [^>]*>                  # start row node
      <w:tc            [^>]*>                  # start cell node
         (?:<w:tcPr> .*? </w:tcPr> (*SKIP) )?  # cell properties
         $regex_paragraph                      # paragraph in cell
      </w:tc>                                  # close cell node
      (?:<w:tc> .*? </w:tc>   (*SKIP) )*       # possibly other cells on the same row
    </w:tr>                                    # close row node
   }sx;

  # assemble template fragments from all runs in the document into a global template text
  my @template_fragments = map {$self->template_fragment_for_run($_)}  @{$self->surgeon->runs};
  my $template_text      = join "", @template_fragments;

  # remove markup for rows around TT2 directives
  $template_text =~ s/$regex_row/$1/g;

  # remove markup for pagraphs around TT2 directives
  $template_text =~ s/$regex_paragraph/$1/g;

  return $template_text;
}






#======================================================================
# PROCESSING THE TEMPLATE
#======================================================================



sub process {
  my ($self, $vars) = @_;

  # invoke the Template Toolkit
  my $template = Template->new($self->template_config)
    or die Template->error(), "\n";
  my $output = "";
  $template->process(\$self->{template_text}, $vars, \$output)
    or die $template->error();

  # insert the generated output into a MsWord document; other zip members are cloned from the original template
  my $new_doc = $self->surgeon->meta->clone_object($self->surgeon);
  $new_doc->contents($output);

  return $new_doc;
}


1;

__END__

=encoding ISO-8859-1

=head1 NAME

MsOffice::Word::Template - process a Word document as a template for the Template Toolkit 

=head1 SYNOPSIS

  my $surgeon  = MsOffice::Word::Surgeon->new($filename);
  my $template = MsOffice::Word::Template->new($surgeon);
  my $docx = $template->process(\%data);
  $docx->save_as($path_for_new_doc);


=head1 DESCRIPTION


=head1 METHODS

=pod


  # use the contents as a template
  $surgeon->reduce_all_noises;
  $surgeon->merge_runs;
  my $template = $surgeon->compile_template(
        highlights => 'yellow',
        engine     => 'Template', # or Mojo::Template
   );
  my $new_doc = $template->process(data => \%some_data_tree, %other_options);
  $new_doc->save_as($new_doc_filename);



THINK
 - how to build paragraphs, list, tables from TT2 directives (not from Word) ?
 - how to include other docs ?
 - TT2 block with several lines, howto ?

