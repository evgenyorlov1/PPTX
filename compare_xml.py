#!/usr/bin/env python

import difflib
import sys
from StringIO import StringIO

from lxml import etree


def xml_assert_equal(expected, actual, max_lines=1000, normalizer=None, ignore_blank_text=True):
    # Transform both documents into an element tree if strings were passed-in
    if isinstance(expected, (str, unicode)):
        expected = etree.parse(StringIO(expected))
    if isinstance(actual, (str, unicode)):
        actual = etree.parse(StringIO(actual))

    # Create a canonical representation of both documents
    if normalizer is not None:
        expected = normalizer(expected)
        actual = normalizer(actual)
    expected = xml_as_canonical_string(expected, remove_blank_text=ignore_blank_text)
    actual = xml_as_canonical_string(actual, remove_blank_text=ignore_blank_text)

    # Then, compute a unified diff from there
    diff = difflib.unified_diff(expected, actual, fromfile='expected.xml', tofile='actual.xml')

    # Print the discrepancies out in unified diff format
    had_differences = False
    line_counter = 0
    for line in diff:
        sys.stdout.write(line)
        had_differences = True
        line_counter += 1
        if line_counter == max_lines:
            sys.stdout.write('<unified diff abbreviated for clarity\'s sake, more lines still to come>')
            break

    if had_differences:
        raise AssertionError('Expected and actual XML seem to differ')


def xml_as_canonical_string(document, schema=None, remove_blank_text=True):
    # Write out the canonical representation as string
    s1 = StringIO()
    if isinstance(document, etree._Element):
        document = etree.ElementTree(document)
    etree.cleanup_namespaces(document)
    document.write_c14n(s1)

    # Make sure it is indented properly before being returned
    s1.seek(0)
    s2 = StringIO()
    parser = etree.XMLParser(
        remove_blank_text=remove_blank_text,
        remove_comments=True,
        ns_clean=True,
        attribute_defaults=True,
    )
    tree = etree.parse(s1, parser=parser)
    tree.write(s2, pretty_print=True, xml_declaration=True, encoding='utf-8')

    # Return the result as an array of strings (one per line)
    s2.seek(0)
    return s2.readlines()


def normalizer(e):
    etree.strip_attributes(e, 'id')
    return e


if __name__ == '__main__':
    with open(sys.argv[1]) as f1:
        with open(sys.argv[2]) as f2:
            xml_assert_equal(f1.read(), f2.read(), normalizer=normalizer)

