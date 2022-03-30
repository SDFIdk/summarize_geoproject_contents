# -*- coding: utf-8 -*-
"""
Utility to summarise the contents of an Esri aprx file.

Created by: Hanne L. Petersen <halpe@sdfe.dk>
Created: 2019(ish)
"""
import sys
import os
from datetime import datetime as dt
import re

from office_utils import OfficeUtils

is_in_arcgis = 'arcgis' in globals()  # this check MUST be done before anything using arcpy is imported!
import arcpy


def set_msg(msg, severity=0):  # placeholder
    print(msg)


def summarise_aprx(aprx_path, outfile_path, check_paths=False):
    t0 = dt.now()
    set_msg("Starting at  {}".format(t0))
    set_msg("  Reading {}...".format(aprx_path))
    aprx = arcpy.mp.ArcGISProject(aprx_path)
    map_lst = aprx.listMaps()

    map_dic = {}
    for map in map_lst:
        set_msg("  Processing map {}...".format(map.name))
        lyr_lst = map.listLayers()
        map_dic[map.name] = get_lyr_stats(lyr_lst, check_paths)

    # Write outputs
    xlsx_path = os.path.splitext(outfile_path)[0] + '.xlsx'  # ensure it's xlsx and not xls
    set_msg("  Writing {}...".format(xlsx_path))
    sheets = []
    stats = []
    for (mapnam, lyr_stats) in map_dic.items():  # unzip dict to lists
        sheets.append(mapnam)
        stats.append(lyr_stats)
    headings = ['Layer', 'Data path', 'Visibility', 'Transparency', 'Max. scale', 'Min. scale',
                'Visualisation attributes', 'Definition query']
    if check_paths:
        headings = headings[:2] + ['Data OK?'] + headings[2:]
    OfficeUtils.multi_lists2xlsx(stats, xlsx_path, headings,
                                 {'sheetname': sheets, 'widths': [40, 60, 10, 10, 10, 10, 40, 50, 20, 50]})
    set_msg("  Output written to {}.".format(xlsx_path))

    set_msg("Python script duration (h:mm:ss.dddd): " + str(dt.now() - t0)[:-2])


def get_lyr_stats(lyr_lst, check_paths=False):
    lyr_stats = []
    for i, lyr in enumerate(lyr_lst):
        try:  # Ignore non-feature layers (e.g. anno classes)
            lyr.isFeatureLayer
        except (NameError, ValueError, AttributeError):
            continue

        try:
            nam = lyr.longName
        except (NameError, ValueError, AttributeError):
            nam = lyr.name

        try:
            visi = lyr.visible
        except (NameError, ValueError, AttributeError):
            visi = 'N/A'

        path_ok = '-'
        try:
            src = lyr.dataSource  # note: annos don't support dataSource and each anno class is considered a layer...
            if check_paths:
                if not i % 5:
                    set_msg("    checked {} layers".format(i))
                if arcpy.Exists(src):
                    path_ok = 'OK'
                else:
                    path_ok = 'missing'
        except (NameError, ValueError, AttributeError):
            if lyr.isGroupLayer:
                src = '[group layer]'
            else:
                src = ''

        try:
            min_scale = lyr.minThreshold
            max_scale = lyr.maxThreshold
            if not min_scale > 1:
                min_scale = '-'
            if not max_scale > 1:
                max_scale = '-'
        except (NameError, ValueError, AttributeError):
            min_scale = '-'
            max_scale = '-'

        try:
            def_qry = lyr.definitionQuery
        except (NameError, ValueError, AttributeError):
            def_qry = ''

        try:
            cim = lyr.getDefinition('V2')  # V2 = cim_version to be used until next major release
            render_flds = cim.renderer.fields
            if render_flds:
                vis_attr = "Standard: [{}]".format(', '.join(cim.renderer.fields))
            else:
                arcade_expr = cim.renderer.valueExpressionInfo.expression
                lst = re.findall('feature\.([a-zA-Z]+)[^a-zA-Z]', arcade_expr)
                vis_attr = "Custom: [{}]".format(', '.join(set(lst)))
        except (NameError, ValueError, AttributeError):
            vis_attr = ''

        try:
            tr = lyr.transparency
        except (NameError, ValueError, AttributeError) as ne:
            tr = 'N/A'

        # # Symbology fields + symb. class name
        # cim = lyr.getDefinition("V2")
        # # cim.renderer.fields - classical symbology - get values too?
        # # cim.renderer.valueExpressionInfo.expression - Arcade Custom
        # try:
        # # if not lyr.isGroupLayer:
        #     set_msg(lyr.name)
        #     if cim.renderer.fields:
        #         sym_fields = cim.renderer.fields
        #         sym_classes = [c.name for c in cim.featureTemplates]
        #     elif cim.renderer.ValueExpressionInfo:
        #         sym_expr = cim.renderer.ValueExpressionInfo.expression
        #         sym_fields = [n[9:] for n in re.findall('\$feature\.[a-zA-Z]+', sym_expr)]
        #         sym_classes = [c.label for c in cim.renderer.groups[0].classes]
        #     else:
        #         sym_fields = []
        #         sym_classes = []
        # except AttributeError:  # e.g. GroupLayer has no .cim
        # # else:
        #     set_msg("Skipping symbology info for {}...".format(lyr.name))
        #     sym_fields = []
        #     sym_classes = []

        # Look for joins - https://gis.stackexchange.com/a/7715
        # join_lst = join_check(lyr)
        # joins = ', '.join(join_lst)

        if check_paths:
            # lyr_stats.append([nam, src, path_ok, str(visi), tr, max_scale, min_scale, def_qry, ';'.join(sym_fields), ';'.join(sym_classes)])  # , joins])
            lyr_stats.append([nam, src, path_ok, str(visi), tr, max_scale, min_scale, vis_attr, def_qry])  # , joins])
        else:
            lyr_stats.append([nam, src, str(visi), tr, max_scale, min_scale, vis_attr, def_qry])  # , joins])

    return lyr_stats


# def join_check(lyr):
#     tbl_set = set()
#     try:
#         fld_lst = arcpy.Describe(lyr).fields
#         for fld in fld_lst:
#             if fld.name.find('.') > -1:
#                 if fld.name.split('.')[0] != lyr.datasetName:
#                     tbl_set.add(fld.name.split('.')[0])
#     except RuntimeError:  # anno layers have a "layer" per anno class, and these don't have fields
#         return []  # False
#     except AttributeError:  # anno layers have a "layer" per anno class, and these don't have fields (yes, both error types will show up, should clarify why)
#         return []  # False
#     return tbl_set


if __name__ == "__main__":
    aprx = r'C:\Temp\SK_tiles.aprx'
    txt = r'C:\Temp\SK_tiles_overview.xlsx'
    verify = True

    if is_in_arcgis:
        aprx = arcpy.GetParameterAsText(0)
        txt = arcpy.GetParameterAsText(1)
        verify = bool(arcpy.GetParameter(2))
    else:
        if len(sys.argv) > 1:
            aprx = sys.argv[1]
        if len(sys.argv) > 2:
            txt = sys.argv[2]

    summarise_aprx(aprx, txt, verify)
    # os.startfile(txt)
