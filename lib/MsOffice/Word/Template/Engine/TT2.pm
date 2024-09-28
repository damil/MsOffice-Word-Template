package MsOffice::Word::Template::Engine::TT2;
use 5.024;
use Moose;
extends 'MsOffice::Word::Template::Engine';

use Template::AutoFilter;  # a subclass of Template that adds automatic html filtering
use Template::Config;      # loaded explicitly so that we can override its $PROVIDER global variable
use MsOffice::Word::Surgeon::Utils qw(encode_entities);
use MsOffice::Word::Template::Engine::TT2::Provider;

use namespace::clean -except => 'meta';

our $VERSION = '2.05';

#======================================================================
# ATTRIBUTES
#======================================================================

has 'start_tag' => (is => 'ro', isa => 'Str',    default  => "[% ");
has 'end_tag'   => (is => 'ro', isa => 'Str',    default  => " %]");
has 'TT2'       => (is => 'ro', isa => 'Template', lazy => 1, builder => "_TT2", init_arg => undef);

#======================================================================
# LAZY ATTRIBUTE CONSTRUCTORS
#======================================================================

sub _TT2 {
  my ($self) = @_;

  my $TT2_args = delete $self->{_constructor_args};

  # inject precompiled blocks into the Template parser
  my $precompiled_blocks = $self->_precompiled_blocks;
  $TT2_args->{BLOCKS}{$_} //= $precompiled_blocks->{$_} for keys %$precompiled_blocks;

  # instantiate the Template object but supply our own Provider subclass. This way of doing is a little
  # bit rude, but so much easier than using the LOAD_TEMPLATES config options ! ... because here the
  # Template Toolkit will take care of supplying all default values to the Provider.
  local $Template::Config::PROVIDER = 'MsOffice::Word::Template::Engine::TT2::Provider';
  my $tt2 = Template::AutoFilter->new($TT2_args);

  return $tt2;
}

#======================================================================
# METHODS
#======================================================================

sub compile_template {
  my ($self, $template_text) = @_;

  return $self->TT2->template(\$template_text);
}


sub process_part {
  my ($self, $part_name, $package_part, $vars) = @_;

  # extend $vars with a pointer to the part object, so that it can be called from
  # the template, for example for replacing an image
  my $extended_vars = {package_part => $package_part, %$vars};

  return $self->process($part_name, $extended_vars);
}


sub process {
  my ($self, $template_name, $vars) = @_;

  # get the compiled template
  my $tmpl = $self->compiled_template->{$template_name}
    or die "don't have a compiled template for '$template_name'";

  # produce the new contents
  my $new_contents  = $self->TT2->context->process($tmpl, $vars);

  return $new_contents;
}


#======================================================================
# PRE-COMPILED BLOCKS THAT CAN BE INVOKED FROM TEMPLATE DIRECTIVES
#======================================================================

# arbitrary value for the first bookmark id. 100 should be enough to be above other
# bookmarks generated by Word itself. A cleaner but more complicated way would be to parse the template
# to find the highest id number really used.
my $first_bookmark_id = 100;

my %barcode_generator = (

  Code128 => sub {
    my ($to_encode, %options) = @_;
    require Barcode::Code128;

    $options{border}    //= 0;
    $options{show_text} //= 0;
    $options{padding}   //= 0;
    my $bc = Barcode::Code128->new;
    $bc->option(%options);
    return $bc->png($to_encode);
  },

  QRCode => sub {
    my ($to_encode, %options) = @_;
    require Image::PNG::QRCode;
    return Image::PNG::QRCode::qrpng(text => $to_encode);
  },

 );



# precompiled blocks as facilities to be used within templates
sub _precompiled_blocks {

  return {

    # a wrapper block for inserting a Word bookmark
    bookmark => sub {
      my $context     = shift;
      my $stash       = $context->stash;

      # assemble xml markup
      my $bookmark_id = $stash->get('global.bookmark_id') || $first_bookmark_id;
      my $name        = fix_bookmark_name($stash->get('name') || 'anonymous_bookmark');

      my $xml         = qq{<w:bookmarkStart w:id="$bookmark_id" w:name="$name"/>}
                      . $stash->get('content') # content of the wrapper
                      . qq{<w:bookmarkEnd w:id="$bookmark_id"/>};

      # next bookmark will need a fresh id
      $stash->set('global.bookmark_id', $bookmark_id+1);

      return $xml;
    },

    # a wrapper block for linking to a bookmark
    link_to_bookmark => sub {
      my $context = shift;
      my $stash   = $context->stash;

      # assemble xml markup
      my $name    = fix_bookmark_name($stash->get('name') || 'anonymous_bookmark');
      my $content = $stash->get('content');
      my $tooltip = $stash->get('tooltip');
      if ($tooltip) {
        encode_entities($tooltip);
        $tooltip = qq{ w:tooltip="$tooltip"};
      }
      my $xml  = qq{<w:hyperlink w:anchor="$name"$tooltip>$content</w:hyperlink>};

      return $xml;
    },

    # a block for generating a Word field. Can also be used as wrapper.
    field => sub {
      my $context = shift;
      my $stash   = $context->stash;
      my $code    = $stash->get('code');         # field code, including possible flags
      my $text    = $stash->get('content');      # initial text content (before updating the field)

      my $xml     = qq{<w:r><w:fldChar w:fldCharType="begin"/></w:r>}
                  . qq{<w:r><w:instrText xml:space="preserve"> $code </w:instrText></w:r>};
      $xml       .= qq{<w:r><w:fldChar w:fldCharType="separate"/></w:r>$text} if $text;
      $xml       .= qq{<w:r><w:fldChar w:fldCharType="end"/></w:r>};

      return $xml;
    },

    # a block for replacing a placeholder image by a generated barcode
    barcode => sub {

      # get parameters from stash
      my $context      = shift;
      my $stash        = $context->stash;
      my $barcode_type = $stash->get('type');         # either 'Code128' or 'QRCode'
      my $package_part = $stash->get('package_part'); # Word::Surgeon::PackagePart
      my $img          = $stash->get('img');          # title of an existing image to replace
      my $to_encode    = $stash->get('content');      # text to be encoded
      my $options      = $stash->get('options') || {};
      $to_encode =~ s(<[^>]+>)()g;

      # generate PNG image
      my $generator = $barcode_generator{$barcode_type}
        or die "unknown type for barcode generator : '$barcode_type'";
      my $png = $generator->($to_encode, %$options);

      # inject image into the .docx file
      $package_part->replace_image($img, $png);

      return "";
    },

  };
}



#======================================================================
# UTILITY ROUTINES (not methods)
#======================================================================


sub fix_bookmark_name {
  my $name = shift;

  # see https://stackoverflow.com/questions/852922/what-are-the-limitations-for-bookmark-names-in-microsoft-word

  $name =~ s/[^\w_]+/_/g;                              # only digits, letters or underscores
  $name =~ s/^(\d)/_$1/;                               # cannot start with a digit
  $name = substr($name, 0, 40) if length($name) > 40;  # max 40 characters long

  return $name;
}


1;

__END__

=encoding ISO-8859-1

=head1 NAME

MsOffice::Word::Template::Engine::TT2 -- Word::Template engine based on the Template Toolkit

=head1 SYNOPSIS

  my $template = MsOffice::Word::Template->new(docx         => $filename
                                               engine_class => 'TT2',
                                               engine_args  => \%args_for_TemplateToolkit,
                                               );

  my $new_doc  = $template->process(\%data);

See the main synopsis in L<MsOffice::Word::Template>.

=head1 DESCRIPTION

Implements a templating engine for L<MsOffice::Word::Template>, based on the
L<Template Toolkit|Template>.

Like in the regular Template Toolkit, directives like C<INSERT>, C<INCLUDE> or C<PROCESS> can be
used to load subtemplates from other MsWord files -- but the C<none> filter must be
added after the directive, as explained in L</ABOUT HTML ENTITIES>.


=head1 AUTHORING NOTES SPECIFIC TO THE TEMPLATE TOOLKIT

This chapter just gives a few hints for authoring Word templates with the
Template Toolkit.

The examples below use [[double square brackets]] to indicate
segments that should be highlighted in B<green> within the Word template.


=head2 Bookmarks

The template processor is instantiated with a predefined wrapper named C<bookmark>
for generating Word bookmarks. Here is an example:

  Here is a paragraph with [[WRAPPER bookmark name="my_bookmark"]]bookmarked text[[END]].

The C<name> argument is automatically truncated to 40 characters, and non-alphanumeric
characters are replaced by underscores, in order to comply with the limitations imposed by Word
for bookmark names.

=head2 Internal hyperlinks

Similarly, there is a predefined wrapper named C<link_to_bookmark> for generating
hyperlinks to bookmarks. Here is an example:

  Click [[WRAPPER link_to_bookmark name="my_bookmark" tooltip="tip top"]]here[[END]].

The C<tooltip> argument is optional.

=head2 Word fields

A predefined block C<field> generates XML markup for Word fields, like for example :

  Today is [[PROCESS field code="DATE \\@ \"h:mm am/pm, dddd, MMMM d\""]]

Beware that quotes or backslashes must be escaped so that the Template Toolkit parser
does not interpret these characters.

The list of Word field codes is documented at 
L<https://support.microsoft.com/en-us/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51>.

When used as a wrapper, the C<field> block generates a Word field with alternative
text content, displayed before the field gets updated. For example :

  [[WRAPPER field code="TOC \o \"1-3\" \h \z \u"]]Table of contents - press F9 to update[[END]]

The same result can also be obtained by using it as a regular block
with a C<content> argument :

  [[PROCESS field code="TOC \o \"1-3\" \h \z \u" content="Table of contents - press F9 to update"]]


=head2 barcodes

The predefined block C<barcode> generates a barcode image to replace
a "placeholder image" already present in the template. This directive can
appear anywhere in the document, it doesn't have to be next to the image location.

  [[ PROCESS barcode type='QRCode' img='img_name' content="some text" ]]
  # or, used as a wrapper
  [[ WRAPPER barcode type='QRCode' img='img_name']]some text[[ END ]]

Parameters to the C<barcode> block are :

=over

=item type

The type of barcode. The currently implemented types are C<Code128> and C<QRCode>.

=item img

The image unique identifier (corresponding to the I<alternative text> of that image
within the template).

=item content

numeric value or text string that will be encoded as barcode

=item options

a hashref of options that will be passed to the barcode generator.
Currently only the C<Code128> generator takes options; these are
described in the L<Barcode::Code128> Perl module. For example :

  [[PROCESS barcode type="Code128" img="my_image" content=123456 options={border=>4, padding=>50}]]

=back

=head2 ABOUT HTML ENTITIES

This module uses L<Template::AutoFilter>, a subclass of L<Template> that
adds automatically an 'html' filter to every templating directive, so that
characters such as C<< '<' >> or C<< '&' >> are automatically encoded as
HTML entities.

If this encoding behaviour is I<not> appropriate, like for example if the
directive is meant to directly produces OOXML markup, then the 'none' filter
must be explicitly mentioned in order to prevent html filtering :

  [[ PROCESS ooxml_producing_block(foo, bar) | none ]]

This is the case in particular when including
XML fragments from other C<.docx> documents through
C<INSERT>, C<INCLUDE> or C<PROCESS> directives : for example

  [[ INCLUDE lib/fragments/some_other_doc.docx | none ]]


=head1 AUTHOR

Laurent Dami, E<lt>dami AT cpan DOT org<gt>

=head1 COPYRIGHT AND LICENSE

Copyright 2020-2024 by Laurent Dami.

This program is free software, you can redistribute it and/or modify it under the terms of the Artistic License version 2.0.
