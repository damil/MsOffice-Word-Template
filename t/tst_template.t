use strict;
use warnings;
use lib "../../MsOffice-Word-Surgeon/lib";
use lib "../lib";
use MsOffice::Word::Surgeon;
use MsOffice::Word::Template;

(my $dir = $0) =~ s[tst_template.t$][];
$dir ||= ".";
my $template_file = "$dir/etc/tst_template.docx";


my $surgeon = MsOffice::Word::Surgeon->new($template_file);
$surgeon->reduce_all_noises;
$surgeon->merge_runs;




my $template = MsOffice::Word::Template->new(surgeon => $surgeon);
my %data = (
  foo => 'FOFOLLE',
  bar => 'WHISKY',
  list => [ {name => 'toto', value => 123},
            {name => 'blublu', value => 456},
            {name => 'zorb', value => 987},
           ],
);
my $new_doc = $template->process(\%data);
$new_doc->save_as("template_result.docx");


