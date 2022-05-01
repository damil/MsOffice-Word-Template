package MsOffice::Word::Template::Engine::TT2;
use 5.024;
use Moose;
use Template::AutoFilter; # a subclass of Template that adds automatic html filtering

extends 'MsOffice::Word::Template::Engine';

# syntactic sugar for attributes
sub has_inner ($@) {my $attr = shift; has($attr => @_, init_arg => undef, lazy => 1, builder => "_$attr")}

use namespace::clean -except => 'meta';

our $VERSION = '1.02';

#======================================================================
# ATTRIBUTES
#======================================================================

has       'start_tag' => (is => 'ro', isa => 'Str',    default  => "[% ");
has       'end_tag'   => (is => 'ro', isa => 'Str',    default  => " %]");
has_inner 'TT2'       => (is => 'ro', isa => 'Template');



#======================================================================
# LAZY ATTRIBUTE CONSTRUCTORS
#======================================================================


sub _TT2 {
  my ($self) = @_;

  my $TT2_args = delete $self->{_constructor_args};

  # inject precompiled blocks into the Template parser
  my $precompiled_blocks = $self->_precompiled_blocks;
  $TT2_args->{BLOCKS}{$_} //= $precompiled_blocks->{$_} for keys %$precompiled_blocks;

  return Template::AutoFilter->new($TT2_args);
}


#======================================================================
# METHODS
#======================================================================

sub compile_template {
  my ($self, $part_name, $template_text) = @_;

  $self->{_compiled_template}{$part_name} = $self->TT2->template(\$template_text);
}


sub process {
  my ($self, $part_name, $package_part, $vars) = @_;

  # get the compiled template
  my $tmpl         = $self->{_compiled_template}{$part_name}
    or die "don't have a compiled template for '$part_name'";

  # extend $vars with a pointer to the part object, so that it can be called from
  # the template, for example for replacing an image
  my $extended_vars = {package_part => $package_part, %$vars};

  # produce the new contents
  my $new_contents  = $self->TT2->context->process($tmpl, $extended_vars);

  return $new_contents;
}


#======================================================================
# PRE-COMPILED BLOCKS THAT CAN BE INVOKED FROM TEMPLATE DIRECTIVES
#======================================================================

# arbitrary value for the first bookmark id. 100 should most often be above other
# bookmarks generated by Word itself. TODO : would be better to find the highest
# id number really used in the template
my $first_bookmark_id = 100;

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
        # TODO: escape quotes
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


    # a block for replacing an image by a new barcode
    barcode => sub {
      require Barcode::Code128;

      my $context      = shift;
      my $stash        = $context->stash;
      my $package_part = $stash->get('package_part'); # Word::Surgeon::PackagePart
      my $img          = $stash->get('img');          # title of an existing image to replace
      my $to_encode    = $stash->get('content');      # text to be encoded
      $to_encode =~ s(<[^>]+>)()g;

      warn "generating barcode for $to_encode\n";

      #cr�e l'image PNG
      my $bc = Barcode::Code128->new;
      $bc->option(border    => 0,
                  show_text => 0,
                  padding   => 0);
      my $png = $bc->png($to_encode);
      $package_part->replace_image($img, $png);
      return "";
    },
  }
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

=head1 DESCRIPTION

Implements a templating engine for L<MsOffice::Word::Template>, based on the
L<Template Toolkit|Template>.
