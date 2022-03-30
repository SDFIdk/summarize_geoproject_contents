# -*- coding: utf-8 -*-
"""
A basic utility to summarise the contents of a map file.
It isn't pretty or polished, and only covers the cases we've needed so far, but mostly does its thing.

INCLUDEs in the map file are currently ignored.

Created by: Hanne L. Petersen <halpe@sdfe.dk>
Created: June 2021
"""
import sys
import os
import logging
import shutil
from datetime import datetime as dt
import collections
import codecs
import mappyfile  # https://pypi.org/project/mappyfile/
import re

from office_utils import OfficeUtils

log_level = logging.DEBUG


def set_msg(msg, severity=0):  # placeholder
    print(msg)


def main(mapfile_src_cp, outfile_path, split_scale_lyrs=False, drop_includes=True):
    set_msg("Processing {} to {}...".format(mapfile_src_cp, outfile_path))
    t0 = dt.now()
    set_msg("Start time: {}".format(t0))

    # Copy input file, so we can kill the bom and any includes
    mapfile_src = mapfile_src_cp
    mapfile_src_cp = '_tmp'.join(os.path.splitext(mapfile_src_cp))
    set_msg("Duplicating map file to {}...".format(mapfile_src_cp))
    if drop_includes:
        with open(mapfile_src, "r") as fin:
            with open(mapfile_src_cp, "w") as fout:
                for line in fin:
                    fout.write(re.sub('(\t|  |\t )INCLUDE', '\\1# INCLUDE', line))
    else:
        shutil.copyfile(mapfile_src, mapfile_src_cp)

    if has_bom(mapfile_src_cp):
        set_msg("Removing BOM...")
        remove_bom_in_place(mapfile_src_cp)

    map_dic = get_lyr_stats(mapfile_src_cp, split_scale_lyrs)

    set_msg("Deleting tmp file...")
    os.remove(mapfile_src_cp)

    # print(map_dic)

    # Write outputs
    xlsx_path = os.path.splitext(outfile_path)[0] + '.xlsx'  # ensure it's xlsx and not xls
    set_msg("  Writing {}...".format(xlsx_path))
    stats = []
    for k in sorted(map_dic.keys()):
        for lyr_info in map_dic[k]:
            stats.append([lyr_info[0]] + [k] + lyr_info[1:])
    # for (mapnam, lyr_stats) in map_dic.items():  # unzip dict to lists
    #     sheets.append(mapnam)
    #     stats.append(lyr_stats)
    # print(stats)
    headings = ['Layer name', 'Layer group', 'Data path (beta)', 'Visibility', 'Max. scale', 'Min. scale', 'Data SQL',
                'WMS Title', 'WMS Layer Group', 'WMS Abstract', 'WMS Group Title', 'WMS Group Abstract']  # , 'Definition query']
    OfficeUtils.multi_lists2xlsx([stats], xlsx_path, headings,
                                 {'widths': [20, 30, 70, 10, 10, 10, 50, 20, 20, 20, 20]})
    set_msg("  Output written to {}.".format(xlsx_path))

    set_msg("Script duration: {}".format(dt.now() - t0))


def get_lyr_stats(mapfile_src, split_scale_lyrs=False):
    mapfile = mappyfile.open(mapfile_src)
    layers = mapfile["layers"]

    groups = collections.defaultdict(list)

    for lyr in layers:
        nam = lyr['name']
        if nam == 'Byomraade':
            print(nam)
            print(lyr['metadata'].keys())

        min_scale = str(lyr['maxscaledenom']) if 'maxscaledenom' in lyr.keys() else 'N/A'
        max_scale = str(lyr['minscaledenom']) if 'minscaledenom' in lyr.keys() else 'N/A'
        meta = lyr['metadata']
        wms_title = str(hack_asciify(lyr['metadata']['wms_title'])) if 'wms_title' in lyr['metadata'].keys() else 'N/A'
        wms_layer_group = str(hack_asciify(lyr['metadata']['wms_layer_group'])) if 'wms_layer_group' in lyr['metadata'].keys() else '-'
        wms_abstract = str(hack_asciify(lyr['metadata']['wms_abstract'])) if 'wms_abstract' in lyr['metadata'].keys() else 'N/A'
        wms_group_title = str(hack_asciify(lyr['metadata']['wms_group_title'])) if 'wms_group_title' in lyr['metadata'].keys() else 'N/A'
        wms_group_abstract = str(hack_asciify(lyr['metadata']['wms_group_abstract'])) if 'wms_group_abstract' in lyr['metadata'].keys() else 'N/A'

        visi = lyr['status'] if 'status' in lyr.keys() else 'N/A'  # ON -> True

        # Do other stuff first - extract token values (key => combined string)
        tokens = {}
        if lyr['scaletokens']:  # Get scaletokens for text replacement
            for t in lyr['scaletokens']:
                # print("{}: {}".format(t['name'], t['values']))
                token_dict = t['values']
                del token_dict['__type__']
                for k, v in token_dict.items():  # map, keep OrderedDict...
                    # print(k)
                    token_dict[k] = strip_from_string(v, ' using')
                long_vals = token_dict.values()
                long_vals_encoded = [encodeIfUnicode(s) for s in long_vals]
                tokens[t['name']] = '[{}]'.format(';'.join(long_vals_encoded))  # I thought py3 did utf, but this fails w ø in scaletoken values
        if 'data' in lyr.keys():
            src_long = lyr['data'][0]
        elif 'connection' in lyr.keys():
            src_long = lyr['connection']
        else:
            src_long = 'N/A'

        src_short = strip_from_string(src_long, ' using')  # src_long[:src_long.lower().index(' using')]
        src_short = src_short.replace('geometri from ', '')

        # Check need for replacement
        do_replace = False
        for rep in tokens.keys():
            if rep in src_short:
                do_replace = True

        grp = lyr['group']
        if not grp:
            grp = '-'  # wms_layer_group
        if not do_replace or not split_scale_lyrs:
            # Replace tokens (if needed)
            for rep in tokens.keys():
                src_short = src_short.replace(rep, tokens[rep])

            groups[grp].append([nam, src_short, visi, max_scale, min_scale, src_long,
                                wms_title, wms_layer_group, wms_abstract, wms_group_title, wms_group_abstract])
        else:
            if lyr['scaletokens']:  # Get scaletokens for text replacement
                for t_grp in lyr['scaletokens']:
                    if t_grp['name'] in src_short:
                        # Traverse with index, to allow peek ahead (or backwards, to avoid peek)
                        keys = t_grp['values'].keys()
                        # for scl, val in six.iteritems(t_grp['values']):
                        for kk in range(len(keys)):
                            scl = keys[kk]
                            try:
                                next_scl = int(keys[kk+1])-1
                            except:
                                next_scl = '-'
                            groups[grp].append([nam, src_short.replace(t_grp['name'], t_grp['values'][scl]),
                                                visi, int(scl), next_scl, src_long,
                                                wms_title, wms_layer_group, wms_abstract, wms_group_title, wms_group_abstract])

    return groups


def hack_asciify(s):
    return encodeIfUnicode(s).replace('æ', 'ae').replace('ø', 'oe').replace('å', 'aa').replace('Æ', 'AE').replace('Ø', 'OE').replace('Å', 'AA')


def strip_from_string(s, search):
    try:
        return s[:s.lower().index(search)]
    except ValueError:
        return s


def encodeIfUnicode(strval):
    """Encode if string is unicode."""
    if sys.version_info.major >= 3:
        return strval
    try:
        if isinstance(strval, unicode):
            return strval.encode('utf8')
    except:
        pass
    return str(strval)


def has_bom(f):
    """From https://gist.github.com/nevill/6a59ad277342bea2f8108cf55a35ba3e"""
    UTF8_BOM = b'\xef\xbb\xbf'
    UTF16_BOM = b'\xff\xfe'
    with open(f, 'rb') as bs:
        s = bs.read(3)
        if s.startswith(UTF8_BOM) or s.startswith(UTF16_BOM):
            return True
    return False


def remove_bom_in_place(path):
    """Removes BOM mark, if it exists, from a file and rewrites it in-place."""
    # From: https://www.stefangordon.com/remove-bom-mark-from-text-files-in-python/
    buffer_size = 4096
    bom_length = len(codecs.BOM_UTF8)

    with open(path, "r+b") as fp:
        chunk = fp.read(buffer_size)
        if chunk.startswith(codecs.BOM_UTF8):
            i = 0
            chunk = chunk[bom_length:]
            while chunk:
                fp.seek(i)
                fp.write(chunk)
                i += len(chunk)
                fp.seek(bom_length, os.SEEK_CUR)
                chunk = fp.read(buffer_size)
            fp.seek(-bom_length, os.SEEK_CUR)
            fp.truncate()


if __name__ == "__main__":
    the_map = r'C:\Temp\my_map.map'
    xlsx = r'C:\Temp\my_map.xlsx'

    if len(sys.argv) < 2 or len(sys.argv) > 4:
        print("Usage: python " + __file__ + " MAPFILE_SRC XLSX_TARG")
    if len(sys.argv) > 1:
        the_map = sys.argv[1]
        xlsx = sys.argv[2]

    main(the_map, xlsx, True)
