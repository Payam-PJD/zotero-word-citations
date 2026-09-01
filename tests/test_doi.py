import unittest

from zotero_word_citations.doi import (
    find_doi_like_values,
    find_doi_tokens,
    parse_doi_groups,
    parse_parenthesized_dois,
)


class ParseDoiTests(unittest.TestCase):
    def test_single_multiple_and_repeated_groups(self):
        text = (
            "One (10.1000/ABC). Two (10.2000/one; 10.3000/TWO), "
            "again (10.1000/abc)."
        )
        groups = parse_parenthesized_dois(text)
        self.assertEqual(3, len(groups))
        self.assertEqual(("10.1000/abc",), groups[0].dois)
        self.assertEqual(("10.2000/one", "10.3000/two"), groups[1].dois)
        self.assertEqual(("10.1000/abc",), groups[2].dois)

    def test_doi_prefixes_are_accepted(self):
        groups = parse_parenthesized_dois(
            "(doi: 10.1000/ABC; https://doi.org/10.2000/xyz)"
        )
        self.assertEqual(("10.1000/abc", "10.2000/xyz"), groups[0].dois)

    def test_ordinary_parentheses_are_not_modified(self):
        groups = parse_parenthesized_dois(
            "Keep (Smith 2020) and (10.1000/abc; not a DOI)."
        )
        self.assertEqual([], groups)

    def test_all_wrapper_and_prefix_forms(self):
        text = (
            "A [10.1000/ONE, https://doi.org/10.2000/two]. "
            "B {doi: 10.3000/three}. "
            r"C \citep{10.4000/four;10.5000/five}."
        )
        groups = parse_doi_groups(text)
        self.assertEqual(["square", "curly", "bibtex"], [g.format for g in groups])
        self.assertEqual(("10.1000/one", "10.2000/two"), groups[0].dois)
        self.assertEqual(("10.3000/three",), groups[1].dois)
        self.assertEqual(("10.4000/four", "10.5000/five"), groups[2].dois)
        self.assertTrue(groups[2].source.startswith(r"\citep{"))

    def test_bare_sequences_merge_on_comma_semicolon_or_hyphen(self):
        text = (
            "10.1000/a-10.2000/b;10.3000/c, https://doi.org/10.4000/d."
        )
        groups = parse_doi_groups(text)
        self.assertEqual(1, len(groups))
        self.assertEqual(
            ("10.1000/a", "10.2000/b", "10.3000/c", "10.4000/d"),
            groups[0].dois,
        )
        self.assertFalse(groups[0].source.endswith("."))

    def test_old_style_suffix_parentheses_and_angle_characters(self):
        text = "(10.1002/(SICI)1099-0844(199912)17:4<290::AID-CBF849>3.0.CO;2-P)"
        groups = parse_doi_groups(text)
        self.assertEqual(1, len(groups))
        self.assertEqual(
            "10.1002/(sici)1099-0844(199912)17:4<290::aid-cbf849>3.0.co;2-p",
            groups[0].dois[0],
        )

    def test_format_filter_and_excluded_active_field_range(self):
        text = "10.1000/bare [https://doi.org/10.2000/url]"
        url_start = text.index("[")
        groups = parse_doi_groups(text, formats={"url"})
        self.assertEqual(1, len(groups))
        self.assertEqual("square", groups[0].format)
        self.assertEqual([], parse_doi_groups(text, excluded_ranges=[(0, len(text))]))
        self.assertEqual(
            [],
            parse_doi_groups(
                text,
                formats={"url"},
                excluded_ranges=[(url_start, len(text))],
            ),
        )

    def test_doi_like_token_is_kept_for_later_metadata_validation(self):
        tokens = find_doi_tokens("unknown 10.9999/not-a-real-record")
        self.assertEqual("10.9999/not-a-real-record", tokens[0].doi)

    def test_literal_citep_placeholder_is_removed_as_one_span(self):
        text = r"Before \citep{10.20944/preprints202311.0688.v1} after"
        groups = parse_doi_groups(text)
        self.assertEqual(1, len(groups))
        self.assertEqual(
            r"\citep{10.20944/preprints202311.0688.v1}", groups[0].source
        )
        self.assertEqual("bibtex", groups[0].format)
        self.assertEqual(("10.20944/preprints202311.0688.v1",), groups[0].dois)

    def test_doi_glued_to_letters_or_digits_is_not_silently_dropped(self):
        text = (
            "dementia10.1007/s11914-023-00847-x, "
            "10.1080/03008207.2020.1682282, "
            "10.1007/s11914-023-00848-w and "
            "Exam710.1161/JAHA.122.026460"
        )
        groups = parse_doi_groups(text)
        self.assertEqual(2, len(groups))
        self.assertEqual(
            (
                "10.1007/s11914-023-00847-x",
                "10.1080/03008207.2020.1682282",
                "10.1007/s11914-023-00848-w",
            ),
            groups[0].dois,
        )
        self.assertEqual(("10.1161/jaha.122.026460",), groups[1].dois)

    def test_space_separated_dois_merge_and_wrapper_is_removed(self):
        text = "Before (10.1000/one 10.2000/two, - 10.3000/three.) after"
        groups = parse_doi_groups(text)
        self.assertEqual(1, len(groups))
        self.assertEqual("(10.1000/one 10.2000/two, - 10.3000/three.)", groups[0].source)
        self.assertEqual(
            ("10.1000/one", "10.2000/two", "10.3000/three"),
            groups[0].dois,
        )

    def test_space_separator_does_not_merge_across_word_paragraphs(self):
        groups = parse_doi_groups("10.1000/one\r10.2000/two\n10.3000/three")
        self.assertEqual(3, len(groups))

    def test_reconciliation_finder_has_no_left_boundary_requirement(self):
        self.assertEqual(
            ["10.1234/abc", "10.5678/def"],
            find_doi_like_values("word10.1234/abc and 710.5678/def."),
        )


if __name__ == "__main__":
    unittest.main()
