package MsOffice::Word::Template::Engine;
use 5.024;
use Moose;
use MooseX::AbstractMethod;
use HTML::Entities                 qw(decode_entities);

# syntactic sugar for attributes
sub has_slot ($@) {my $attr = shift; has($attr => @_, is => 'bare', init_arg => undef)}

use namespace::clean -except => 'meta';

our $VERSION = '2.0';

#======================================================================
# ATTRIBUTES
#======================================================================

has_slot '_constructor_args'  => (isa => 'HashRef');
has_slot '_compiled_template' => (isa => 'HashRef');

#======================================================================
# ABSTRACT METHODS -- to be defined in subclasses
#======================================================================

abstract 'start_tag';
abstract 'end_tag';
abstract 'compile_template';
abstract 'process_part';
abstract 'process';

#======================================================================
# GLOBALS
#======================================================================

my $XML_COMMENT_FOR_MARKING_DIRECTIVES = '<!--TEMPLATE_DIRECTIVE_ABOVE-->';

#======================================================================
# INSTANCE CONSTRUCTION
#======================================================================


sub BUILD {
  my ($self, $args) = @_;
  $self->{_constructor_args} = $args; # stored to be available for lazy attr constructors
}


#======================================================================
# COMPILING INNER TEMPLATES
#======================================================================


sub _compile_templates {
  my ($self, $word_template) = @_;

  my $data_color    = $word_template->data_color;
  my $control_color = $word_template->control_color;
  my $surgeon       = $word_template->surgeon;

  # compile regexes to be tried on each run in the document
  my @xml_regexes = $self->_xml_regexes;

  # build a compiled template for each document part
  foreach my $part_name ($word_template->part_names->@*) {
    my $part = $surgeon->part($part_name);

    # assemble template fragments from all runs in the part into a global template text
    $part->cleanup_XML;
    my @template_fragments = map {$self->_template_fragment_for_run($_, $data_color, $control_color)}
                                 $part->runs->@*;
    my $template_text      = join "", @template_fragments;

    # remove markup around directives, successively for table rows, for paragraphs, and finally
    # for remaining directives embedded within text runs.
    $template_text =~ s/$_/$1/g foreach @xml_regexes;

    # compile and store the template
    $self->{_compiled_template}{$part_name} = $self->compile_template($template_text);
  }

  # build a compiled template for each property file (core.xml, app.xml, custom.xml)
  foreach my $property_file ($word_template->property_files->@*) {
    if ($surgeon->zip->memberNamed($property_file)) {
      my $xml = $surgeon->xml_member($property_file);
      $self->{_compiled_template}{$property_file} = $self->compile_template($xml);
    }
  }
}


sub _template_fragment_for_run { # given an instance of Surgeon::Run, build a template fragment
   my ($self, $run, $data_color, $control_color) = @_;

  my $props         = $run->props;

  # if this run is highlighted in data or control color, it must be translated into a template directive
  if ($props =~ s{<w:highlight w:val="($data_color|$control_color)"/>}{}) {
    my $color       = $1;
    my $xml         = $run->xml_before;

    # re-build the run, removing the highlight, and adding the start/end tags for the template engine
    my $inner_texts = $run->inner_texts;
    if (@$inner_texts) {
      $xml .= "<w:r>";                                                # opening XML tag for run node
      $xml .= "<w:rPr>" . $props . "</w:rPr>" if $props;              # optional run properties
      $xml .= "<w:t>";                                                # opening XML tag for text node
      $xml .= $self->start_tag;                                       # start a template directive
      foreach my $inner_text (@$inner_texts) {                        # loop over text nodes
        my $txt = decode_entities($inner_text->literal_text);         # just take inner literal text
        $xml .= $txt . "\n";
        # NOTE : adding "\n" because the template parser may need them for identifying end of comments
      }

      $xml .= $self->end_tag;                                         # end of template directive
      $xml .= $XML_COMMENT_FOR_MARKING_DIRECTIVES
                                         if $color eq $control_color; # XML comment for marking
      $xml .= "</w:t>";                                               # closing XML tag for text node
      $xml .= "</w:r>";                                               # closing XML tag for run node
    }

    return $xml;
  }

  # otherwise this run is just regular MsWord content
  else {
    return $run->as_xml;
  }
}


sub _xml_regexes {
  my ($self) = @_;

  # start and end character sequences for a template fragment
  my $rx_start = quotemeta $self->start_tag;
  my $rx_end   = quotemeta $self->end_tag;

  # Regexes for extracting template directives within the XML.
  # Such directives are identified through a specific XML comment -- this comment is
  # inserted by method "template_fragment_for_run()" below.
  # The (*SKIP) instructions are used to avoid backtracking after a
  # closing tag for the subexpression has been found. Otherwise the
  # .*? inside could possibly match across boundaries of the current
  # XML node, we don't want that.

  # regex for matching directives to be treated outside the text flow.
  my $rx_outside_text_flow = qr{
      <w:r\b           [^>]*>                  # start run node
        (?: <w:rPr> .*? </w:rPr>   (*SKIP) )?  # optional run properties
        <w:t\b         [^>]*>                  # start text node
          ($rx_start .*? $rx_end)  (*SKIP)     # template directive
          $XML_COMMENT_FOR_MARKING_DIRECTIVES  # specific XML comment
        </w:t>                                 # close text node
      </w:r>                                   # close run node
   }sx;

  # regex for matching paragraphs that contain only a directive
  my $rx_paragraph = qr{
    <w:p\b             [^>]*>                  # start paragraph node
      (?: <w:pPr> .*? </w:pPr>     (*SKIP) )?  # optional paragraph properties
      $rx_outside_text_flow
    </w:p>                                     # close paragraph node
   }sx;

  # regex for matching table rows that contain only a directive in the first cell
  my $rx_row = qr{
    <w:tr\b            [^>]*>                  # start row node
      <w:tc\b          [^>]*>                  # start cell node
         (?:<w:tcPr> .*? </w:tcPr> (*SKIP) )?  # cell properties
         $rx_paragraph                         # paragraph in cell
      </w:tc>                                  # close cell node
      (?:<w:tc> .*? </w:tc>        (*SKIP) )*  # ignore other cells on the same row
    </w:tr>                                    # close row node
   }sx;

  return ($rx_row, $rx_paragraph, $rx_outside_text_flow);
  # Note : the order is important -- the most specific regex is tried first, the least specific is tried last
}


1;

__END__

=encoding ISO-8859-1

=head1 NAME

MsOffice::Word::Template::Engine -- abstract class for template engines

=head1 DESCRIPTION

This abstract class encapsulates functionalities common to all templating engines.
Concrete classes such as L<MsOffice::Word::Template::Engine::TT2> inherit from the
present class.

Templating engines encapsulate internal implementation algorithms; they are not meant to
be called from external clients. Methods documented below are just to explain the internal
architecture.

=head1 METHODS

=head2 _compile_templates

  $self->_compile_templates($word_template);

Calls the subclass's concrete method C<compile_template> on each document part.

=head2 _template_fragment_for_run

Translates a given text run into a fragment suitable to be processed by the template compiler.

=head2 xml_regexes

Compiles the regexes to be tried on each text run.








