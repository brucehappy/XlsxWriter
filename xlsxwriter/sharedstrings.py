###############################################################################
#
# SharedStrings - A class for writing the Excel XLSX sharedStrings file.
#
# Copyright 2013-2019, John McNamara, jmcnamara@cpan.org
#

# Standard packages.
import re
import abc

# Package imports.
from . import xmlwriter
from .utility import escape_string


class SharedStrings(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX sharedStrings file.

    """

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def __init__(self):
        """
        Constructor.

        """

        super(SharedStrings, self).__init__()

        self.string_table = None

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self):
        # Assemble and write the XML file.

        # Write the XML declaration.
        self._xml_declaration()

        # Write the sst element.
        self._write_sst()

        # Write the sst strings.
        self._write_sst_strings()

        # Close the sst tag.
        self._xml_end_tag('sst')

        # Close the file.
        self._xml_close()

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_sst(self):
        # Write the <sst> element.
        xmlns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

        attributes = [
            ('xmlns', xmlns),
            ('count', self.string_table.count),
            ('uniqueCount', self.string_table.unique_count),
        ]

        self._xml_start_tag('sst', attributes)

    def _write_sst_strings(self):
        # Write the sst string elements.

        for string in self.string_table.get_strings():
            self._write_si(string)

    def _write_si(self, string):
        # Write the <si> element.
        attributes = []

        # Escape the string.
        string = escape_string(string)

        # Add attribute to preserve leading or trailing whitespace.
        if re.search(r'^\s', string) or re.search(r'\s$', string):
            attributes.append(('xml:space', 'preserve'))

        # Write any rich strings without further tags.
        if re.search('^<r>', string) and re.search('</r>$', string):
            self._xml_rich_si_element(string)
        else:
            self._xml_si_element(string, attributes)


# A metadata class to store Excel strings between worksheets.
class AbstractSharedStringTable:
    __metaclass__ = abc.ABCMeta

    @abc.abstractproperty
    def supports_constant_memory(self):
        pass

    @abc.abstractproperty
    def unique_count(self):
        pass

    @abc.abstractproperty
    def count(self):
        pass

    @abc.abstractmethod
    def get_index(self, string):
        pass

    @abc.abstractmethod
    def get_string(self, index):
        pass

    @abc.abstractmethod
    def get_strings(self):
        pass


class SharedStringTable(object):
    """ A class to track Excel shared strings between worksheets. """
    def __init__(self):
        self.count = 0
        from bidict import OrderedBidict
        self._strings = OrderedBidict()
        self.supports_constant_memory = False

    @property
    def unique_count(self):
        return len(self._strings)

    def get_index(self, string):
        """" Get the index of the string in the Shared String table. """
        if string not in self._strings:
            # String isn't already stored in the table so add it.
            index = self.unique_count
            self._strings[string] = index
            self.count += 1
            return index
        else:
            # String exists in the table.
            index = self._strings[string]
            self.count += 1
            return index

    def get_string(self, index):
        """" Get a shared string from the index. """
        return self._strings.inverse[index]

    def get_strings(self):
        """" Return the sorted string iterator. """
        return self._strings.iterkeys()


AbstractSharedStringTable.register(SharedStringTable)
