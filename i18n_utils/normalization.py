import six

import polib


DEFAULT_WRAPPING_WIDTH = 78


def one_occurrence_per_line(entry_str):
    entry_lines = []

    for line in entry_str.split('\n'):
        if line.startswith('#: '):
            line = line[3:]
            for occurrence in line.split(' '):
                entry_lines.append('#: {}'.format(occurrence))
        else:
            entry_lines.append(line)

    return u'\n'.join(entry_lines)


def entry_to_unicode(entry, wrapwidth):
    return one_occurrence_per_line(entry.__unicode__(wrapwidth))


class NormalizedPOFile(polib.POFile):
    def __unicode__(self):
        """
        Returns the normalized unicode representation of the file.
        """
        meta, headers = '', self.header.split('\n')
        for header in headers:
            if header[:1] in [',', ':']:
                meta += u'#%s\n' % header
            else:
                meta += u'# %s\n' % header

        ret = []
        entries = [self.metadata_as_entry()]
        entries += sorted(self, key=lambda e: e.msgid)
        for entry in entries:
            ret.append(entry_to_unicode(entry, self.wrapwidth))
        ret = u'\n'.join(ret)
        assert isinstance(ret, six.text_type)
        return meta + ret
