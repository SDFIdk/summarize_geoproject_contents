# -*- coding: utf-8 -*-
"""
Utility to summarise the contents of an Esri mxd file.

Created by: Hanne L. Petersen <halpe@sdfe.dk>
Created: 2019(ish)
"""
import sys

from office_utils import OfficeUtils

is_in_arcgis = 'arcgis' in globals()  # this check MUST be done before anything using arcpy is imported!
import arcpy


def set_msg(msg, severity=0):  # placeholder
    print(msg)


def summarise_mxd(mxd_path, outfile_path=None, check_paths=False):
    mxd = arcpy.mapping.MapDocument(mxd_path)
    lyr_lst = arcpy.mapping.ListLayers(mxd)
    set_msg("  Reading {}...".format(mxd_path))

    lyr_stats = get_lyr_stats(lyr_lst, check_paths)

    # Write outputs
    if outfile_path:
        if outfile_path.split('.')[1] in ('xls', 'xlsx'):   # Write xlsx
            xlsx_path = outfile_path.split('.')[0] + '.xlsx'  # ensure it's xlsx and not xls
            headings = ['Layer', 'Data path', 'Visibility', 'Transparency', 'Max. scale',
                        'Min. scale', 'Definition query', 'Joins']
            if check_paths:
                headings = headings[:2] + ['Data OK?'] + headings[2:]
            OfficeUtils.lists2xlsx(lyr_stats, xlsx_path, headings)
            set_msg("  Output written to {}.".format(xlsx_path))

        else:  # Write txt
            with open(outfile_path, 'w') as wf:
                for l in lyr_stats:
                    wf.write(encode_if_unicode('\t'.join(l) + '\n'))
            set_msg("  Output written to {}.".format(outfile_path))

    # Output to screen
    else:
        set_msg(['\n'.join(['\t'.join(l) for l in lyr_stats])])


def get_lyr_stats(lyr_lst, check_paths):
    lyr_stats = []
    for lyr in lyr_lst:
        nam = encode_if_unicode(lyr.longName)
        try:
            nam = nam.decode('utf8')
        except AttributeError:
            pass

        visi = lyr.visible

        try:
            src = lyr.dataSource  # annos don't support dataSource...
            if check_paths:
                if arcpy.Exists(src):
                    path_ok = 'OK'
                else:
                    path_ok = 'missing'
        except (NameError, ValueError):
            if lyr.isGroupLayer:
                src = '[group layer]'
            else:
                src = ''

        min_scale = lyr.minScale
        max_scale = lyr.maxScale
        if not min_scale > 1:
            min_scale = '-'
        if not max_scale > 1:
            max_scale = '-'

        try:
            def_qry = lyr.definitionQuery
        except (NameError, ValueError):
            def_qry = ''

        try:
            tr = lyr.transparency
        except (NameError, ValueError) as ne:  # NameError?
            tr = 'N/A'

        # Look for joins - https://gis.stackexchange.com/a/7715
        join_lst = join_check(lyr)
        joins = ', '.join(join_lst)

        if check_paths:
            lyr_stats.append([nam, src, path_ok, str(visi), tr, max_scale, min_scale, def_qry, joins])
        else:
            lyr_stats.append([nam, src, str(visi), tr, max_scale, min_scale, def_qry, joins])

    return lyr_stats


def join_check(lyr):
    tbl_set = set()
    try:
        fld_lst = arcpy.Describe(lyr).fields
        for fld in fld_lst:
            if fld.name.find('.') > -1:
                if fld.name.split('.')[0] != lyr.datasetName:
                    tbl_set.add(fld.name.split('.')[0])
    except (RuntimeError, AttributeError):  # anno layers have a "layer" per anno class, and these don't have fields
        return []  # False
    return tbl_set


def encode_if_unicode(strval):
    """Encode if string is unicode."""
    if isinstance(strval, unicode):
        return strval.encode('utf8')
    return str(strval)


if __name__ == "__main__":
    mxd = r'C:\Temp\my_project.mxd'
    txt = r'C:\Temp\mxd_overview.xlsx'

    if is_in_arcgis:
        mxd = arcpy.GetParameterAsText(0)
        txt = arcpy.GetParameterAsText(1)
    else:
        if len(sys.argv) > 1:
            mxd = sys.argv[1]
        if len(sys.argv) > 2:
            txt = sys.argv[2]

    summarise_mxd(mxd, txt)
