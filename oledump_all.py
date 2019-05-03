#!/usr/bin/env python

""" Find all files embedded in ole, dump them with correct extension

Main work is done by functions called within original oledump.py.

TODO:
* more robust open, that understands more different file types (e.g. open xml)
* add some relevant options from oledump, like decoders or decryption

License:
Same as oledump.py. Use at your own risk.
"""

from __future__ import print_function

import sys
import os
import os.path as op
from zipfile import is_zipfile, ZipFile
from argparse import ArgumentParser, ArgumentTypeError
from traceback import print_exc

try:
    import olefile
except ImportError:
    # maybe oletools are installed, try importing from them
    olefile = None

if olefile is None:
    try:
        from oletools.thirdparty import olefile
    except ImportError:
        raise ImportError('need olefile or oletools installed!')

    # if this worked, then we can do a little hack to import olefile in oledump
    THIRDPARTY_DIR = op.dirname(op.dirname(olefile.__file__))
    if THIRDPARTY_DIR not in sys.path:
        sys.path.append(THIRDPARTY_DIR)
    del THIRDPARTY_DIR

import oledump.oledump as dumper

__description__ = 'Automate oledump.py to dump all embedded files'
__author__ = 'Christian Herdtweck'
__version__ = '0.0.30'
__date__ = '2017/11/10'


# return values from main
RETURN_NO_EXTRACT = 0      # all input files were clean
RETURN_DID_EXTRACT = 1     # did extract files
RETURN_ARGUMENT_ERR = 2    # reserved for parse_args
RETURN_OPEN_FAIL = 3       # failed to open a file
RETURN_STREAM_FAIL = 4     # failed to open an OLE stream

MAX_SIZE = 100*1024*1024   # 100 MB


def existing_file(filename):
    """ called by argument parser to see whether given file exists """
    if not op.isfile(filename):
        raise ArgumentTypeError('{0} is not a file.'.format(filename))
    return filename


def parse_args(cmd_line_args=None):
    """ parse command line arguments (sys.argv by default) """
    parser = ArgumentParser(description='Extract files embedded in OLE')
    parser.add_argument('-d', '--target-dir', type=str, default='.',
                        help='Directory to extract files to. File names are '
                             '0.ext, 1.ext ... . Default: current working dir')
    parser.add_argument('-v', '--verbose', action='store_true',
                        help='verbose mode, not implemented yet')
    parser.add_argument('-i', '--more-input', metavar='FILE',
                        type=existing_file,
                        help='Additional file to parse (same as positional '
                             'arguments)')
    parser.add_argument('input_files', metavar='FILE', nargs='*',
                        type=existing_file,
                        help='Office files to parse (same as -i)')
    args = parser.parse_args(cmd_line_args)

    # combine arguments with -i (compatibility with ripOLE)
    if args.more_input:
        args.input_files += [args.more_input,]
    if not args.input_files:
        parser.error('No input given (use -i and/or positional argument[s])')

    return args


def open_file(filename):
    """ try to open somehow as zip or ole or so; raise exception if fail """
    try:
        if olefile.isOleFile(filename):
            print('is ole file: ' + filename)
            # todo: try ppt_parser first
            yield olefile.OleFileIO(filename)
        elif is_zipfile(filename):
            print('is zip file: ' + filename)
            zipper = ZipFile(filename, 'r')
            for subfile in zipper.namelist():
                head = b''
                try:
                    head = zipper.open(subfile).read(len(olefile.MAGIC))
                except RuntimeError:
                    print('zip is encrypted: ' + filename)  # todo: passwords?!
                    yield None

                if head == olefile.MAGIC:
                    print('  unzipping ole: ' + subfile)
                    yield olefile.OleFileIO(zipper.open(subfile)
                                            .read(MAX_SIZE))
                else:
                    pass  # print('unzip skip: ' + subfile)
        else:  # todo: add more file types
            print('open failed: ' + filename)
            yield None   # --> leads to non-0 return code
    except Exception:
        print_exc()
        yield None   # --> leads to non-0 return code


def ole_iter_streams(ole):
    """ Iterate over all streams (including orphans) in given OleFileIO

    modified copy of oledump.OLEGetStreams, only it yields and does not
    read() streams
    """
    for fname in ole.listdir(streams=True, storages=False):
        name_to_yield = fname
        if not isinstance(fname, (list, tuple)):
            name_to_yield = (fname, )   # ensure it is a tuple
        yield False, name_to_yield, ole.get_type(fname), ole.openstream(fname)
    for sid in range(len(ole.direntries)):
        entry = ole.direntries[sid]
        if entry is None:
            entry = ole._load_direntry(sid)
            if entry.entry_type == olefile.STGTY_STREAM:
                name_to_yield = entry.name
                if not isinstance(name_to_yield, (list, tuple)):
                    name_to_yield = (entry.name, )  # ensure it is a tuple
                yield True, name_to_yield, entry.entry_type, \
                      ole._open(entry.isectStart, entry.size)
    # TODO: can probably skip the first loop since 2nd finds the same
    # entries again with entry != None (see oletools/olevba.py)


def main(cmd_line_args=None):
    """ Main function, called when running file as script

    returns one of the RETURN_* values

    see module doc for more info
    """
    args = parse_args(cmd_line_args)   # does a sys.exit(2) if parsing fails

    output_count = 0
    return_value = RETURN_NO_EXTRACT

    # loop over file name arguments
    for filename in args.input_files:

        # loop over files found within filename
        for ole in open_file(filename):
            if ole is None:
                return_value = max(return_value, RETURN_OPEN_FAIL)
                continue

            # loop over streams within file
            for is_orphan, st_path, _, stream in ole_iter_streams(ole):
                print('    Checking stream "{0}"{1}'
                      .format('/'.join(st_path),
                              ' (orphan)' if is_orphan else ''))

                # read complete stream into memory
                data = stream.read(MAX_SIZE)   # for dumper.*
                # TODO: actually only need first bytes and the size until end

                # check if this is an embedded file
                if not dumper.OLE10HeaderPresent(data):
                    # print('      not an embedded file - skip')
                    continue

                # get filename options
                embedded_data = dumper.ExtractOle10Native(data)
                if not embedded_data:
                    continue
                fn1, fn2, fn3, contents = embedded_data
                del data   # clear memory
                filenames = (fn1, fn2, fn3)
                if os.name in ('posix', 'mac'):  # convert c:\a.ext --> c/a.ext
                    filenames = [fn.replace('\\', '/').replace(':', '/')
                                 .replace('//', '/') for fn in filenames]
                print('      filenames: {0}'.format(filenames))

                # get extension
                extensions = [op.splitext(filename)[1].strip()
                              for filename in filenames]
                extensions = [ext for ext in extensions if ext]
                # print('      extensions: {0}'.format(extensions))
                if not extensions:
                    # print('      no extension found, use empty')
                    extension = ''
                elif all(ext == extensions[0] for ext in extensions[1:]):
                    # all extensions are the same
                    # print('      all extension are the same')
                    extension = extensions[0]
                else:
                    # print('      multiple extensions, use first')
                    extension = extensions[0]

                # dump
                # todo: add decode/decrypt/... functionality here?
                name = op.join(args.target_dir,
                               'oledump{0}{1}'.format(output_count, extension))
                print('      dumping to ' + name)
                with open(name, 'wb') as writer:
                    writer.write(contents)

                output_count += 1

    if output_count:
        return_value = max(return_value, RETURN_DID_EXTRACT)

    return return_value


if __name__ == '__main__':
    sys.exit(main())
