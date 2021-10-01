# -*- coding: cp1252 -*-

'''

    dossierspot
    ===========

        Construit un dossier spot standard à partir d'une liste de depcoms

        Le rapport est exporté à l'emplacement du fichier depcoms.
        les fichiers intermédiaires sont classés dans des sous-dossiers
        par extention.
        Le nommage est défini par le nom du dossier parent

    usage:
    ------

    >>> import dossierspot as spot
    >>> rap = spot.Report(r'/path/to/depcoms.txt')
    >>> rap.process()
    >>>
'''

# constants
# =========
MOS_PATH = r'\\srvdata\3d\Transferts\multidate_MOS2010_2020.gdb\multidate_MOS2010_2020'
TPL_PATH = r'J:\Etudes\laufma\Python26\site-packages\mezcal\templates\pandemic_edition'
DEPCOM_FLD = 'code_insee'
COD10_NIV2 = 'cod_10niv2'
COD20_NIV2 = 'cod_20niv2'

import sys
import os
import shutil
import arcpy
import win32com.client as win32


def naming(path_to_depcoms):
    """
        retrieves naming from path to depcoms (parent folder)
        :param path_to_depcoms: path to depcoms.txt file
        :type path_to_depcoms': unicode
        :return: base name (parent folder name) for entire report
        :rtype: unicode
    """
    return os.path.basename(os.path.dirname(path_to_depcoms))

class Depcoms(object):
    """
        reads depcoms file
    """
    
    def __init__(self, depcoms_path):
        """
            initiates Depcoms object
            :param depcoms_path: path to depcoms file
            :type depcoms_path: unicode
            :return: Depcoms object containing depcoms as list and string
            :rtype: Depcoms instance
        """
        # depcoms path as archive
        self.path = depcoms_path
        # depcoms as a python list
        self.lst = self.read_depcoms()
        # depcoms as a string  type "'69123', '69381', '69382', '69383'"
        self.strg = "'{0}'".format("', '".join(self.lst))

    def read_depcoms(self):
        """
            utility: reads depcom file
            :return: depcoms as python list
            :rtype: list of unicodes
        """
        with open(self.path) as f:
            return f.read().splitlines()
        

class Page(object):
    """
        defines pages
    """

    def __init__(self, source, depcoms, template):
        """
            initiates a new page object
            :param source: path to source fc
            :param depcoms: depcoms list ['69381', '69382']
            :param template: path to template file
            :type source: unicode
            :type depcoms: list of unicodes
            :type template: unicode
            :return: page including pathes and process methods
            :rtype: Page instance
        """
        # path to source fc
        self.source = source
        # backup depcoms
        self.depcoms = depcoms
        # full path to template
        self.template = template
        # titling
        self.title = naming(depcoms.path)
        # st / evo / ortho / data
        self.type = os.path.basename(template).split('_')[0]
        # path to output file
        self.path = self.get_path()
        # path to pdf
        self.pdf = self.get_path(pdf=True)
        # naming ('st_00', ..., 'st_15', 'data'...)
        self.name = '_'.join(os.path.basename(self.path).split('_')[:-1])
        # processed ?
        self.processed = False

    def get_path(self, pdf=False):
        """
            builds output path
            :param pdf: switches to pdf path
            :type pdf: boolean
            :return: path to output file or pdf
            :rtype: unicode
        """
        # root folder comes from depcoms.txt
        root, _ = os.path.split(self.depcoms.path)
        # filename and ext come from template
        _, basename = os.path.split(self.template)
        # replace naming
        basename = basename.replace(u'tpl', self.title)
        filename, ext = os.path.splitext(basename)
        # if path_type is pdf: overwrite extension
        if pdf:
            return os.path.join(root, u'pdf', u'pages', filename+u'.pdf')
        else:
            return os.path.join(root, ext[1:], filename+ext)

    def process_page(self):
        """
            switcher to tune_* functions: see below
        """
        # make sure path exists
        if not os.path.exists(os.path.dirname(self.path)):
            os.makedirs(os.path.dirname(self.path))
        # make sure pdf path exists
        if not os.path.exists(os.path.dirname(self.pdf)):
            os.makedirs(os.path.dirname(self.pdf))
        # copy template
        shutil.copy(self.template, self.path)
        # switch to processor functions by type
        {'st': self.tune_st,                    # millesime maps
         'evo': self.tune_evo,                  # evo maps
         'ortho': self.tune_ortho,              # ortho maps
         'data': self.tune_data}[self.type]()   # xlsx files

    def tune_st(self):
        """
            tune st mxd
        """
        # access mxd
        mxd = arcpy.mapping.MapDocument(self.path)
        # set title
        mxd.title = self.title
        # set spot layer definition query
        styr_lyr = [
            lyr for lyr in arcpy.mapping.ListLayers(mxd)
            if lyr.name.split()[0] == 'mos'][0]
        styr_lyr.definitionQuery = u'%s in (%s)' % (DEPCOM_FLD, self.depcoms.strg)
        # set mapper zoom (110%)
        styr_extent = styr_lyr.getExtent()
        mapper = [
            df for df in arcpy.mapping.ListDataFrames(mxd)
            if df.name == 'carte'][0]
        mapper.extent = styr_extent
        mapper.scale = mapper.scale * 1.1
        # set cartogram definition queries
        perimfill_lyr = [
            lyr for lyr in arcpy.mapping.ListLayers(mxd)
            if lyr.name == 'perim_fill'][0]
        perimfill_lyr.definitionQuery = u'%s in (%s)' % (DEPCOM_FLD, self.depcoms.strg)
        # save mxd and lyr, del mxd
        mxd.save()
        lyr_path = os.path.join(
            os.path.split(self.depcoms.path)[0],
            'lyr', os.path.split(self.path)[1][:-4] + '.lyr')
        if not os.path.exists(os.path.dirname(lyr_path)):
            os.makedirs(os.path.dirname(lyr_path))
        styr_lyr.saveACopy(lyr_path)
        arcpy.mapping.ExportToPDF(
            mxd, self.pdf,
            resolution=300,
            image_quality = 'BEST',
            compress_vectors = False,
            image_compression = 'NONE')
        del mxd
        # flag self as processed
        self.processed = True

    def tune_evo(self):
        """
            tune evo mxd
        """
        # get mxd, year of origin, year of end
        yro, yre = os.path.basename(self.path).split('_')[1:-1]
        mxd = arcpy.mapping.MapDocument(self.path)
        # set title
        mxd.title = self.title
        # set mapper definition queries
        evo_lyr = [
            lyr for lyr in arcpy.mapping.ListLayers(mxd)
            if lyr.name.split()[0] == 'Evolution'][0]
        styro_lyr = [
            lyr for lyr in arcpy.mapping.ListLayers(mxd)
            if lyr.name.split()[0] == 'Occupation'][0]
        evo_lyr.definitionQuery = (
            u'cod_{yro}niv2 > 40 and cod_{yre}niv2 < 40'
            u'and {depcom_fld} in ({depcoms})').format(
                depcoms=self.depcoms.strg ,
                yro=yro, yre=yre,
                depcom_fld=DEPCOM_FLD)
        styro_lyr.definitionQuery = u'{depcom_fld} in ({depcoms})'.format(
            depcoms=self.depcoms.strg,
            depcom_fld=DEPCOM_FLD)
        # set mapper zoom
        styro_extent = styro_lyr.getExtent()
        mapper = [
            df for df in arcpy.mapping.ListDataFrames(mxd)
            if df.name == 'carte'][0]
        mapper.extent = styro_extent
        mapper.scale = mapper.scale * 1.1
        # set cartogram definition queries
        perimfill_lyr = [
            lyr for lyr in arcpy.mapping.ListLayers(mxd)
            if lyr.name == 'perim_fill'][0]
        perimfill_lyr.definitionQuery = u'CODE_INSEE in ({0})'.format(self.depcoms.strg)
        # save and del
        mxd.save()
        arcpy.mapping.ExportToPDF(
            mxd, self.pdf,
            resolution=300,
            image_quality = 'BEST',
            compress_vectors = False,
            image_compression = 'NONE')
        del mxd
        # flag self as processed
        self.processed = True

    def tune_ortho(self):
        """
            tune ortho mxd
        """
        # get mxd, set pdf path
        mxd = arcpy.mapping.MapDocument(self.path)
        # set title
        mxd.title = self.title
        #   - set mapper definition query
        orthozoom_lyr = [
            lyr for lyr in arcpy.mapping.ListLayers(mxd)
            if lyr.name == 'zoom'][0]
        orthomask_lyr = [
            lyr for lyr in arcpy.mapping.ListLayers(mxd)
            if lyr.name == 'mask'][0]
        orthozoom_lyr.definitionQuery = u'CODE_INSEE in ({0})'.format(self.depcoms.strg)
        orthomask_lyr.definitionQuery = (
            u'not CODE_INSEE like \'6938%\' '
            u'and not CODE_INSEE in ({0})').format(self.depcoms.strg)
        # set mapper zoom
        zoom_extent = orthozoom_lyr.getExtent()
        mapper = [
            df for df in arcpy.mapping.ListDataFrames(mxd)
            if df.name == 'carte'][0]
        mapper.extent = zoom_extent
        mapper.scale = mapper.scale * 1.1
        # set cartogram definition queries
        perimfill_lyr = [
            lyr for lyr in arcpy.mapping.ListLayers(mxd)
            if lyr.name == 'perim_fill'][0]
        perimfill_lyr.definitionQuery = u'CODE_INSEE in ({0})'.format(self.depcoms.strg)
        # save mxd, del
        mxd.save()
        arcpy.mapping.ExportToPDF(
            mxd, self.pdf,
            resolution=300,
            image_quality = 'BEST',
            compress_vectors = False,
            image_compression = 'NONE')
        del mxd
        # flag self as processed
        self.processed = True

    def tune_data(self):
        """
            tune xlsx
        """
        # extract data from lyr file
        # flds = [
        #     'depcom_19',
        #     'ST_00', 'ST_05', 'ST_10', 'ST_15',
        #     'ST5_05', 'ST5_10', 'ST5_15',
        #     'SHAPE@AREA']

        # new tmpl as of 20160526
        flds = [
            COD10_NIV2, COD20_NIV2, 'surf_m2', DEPCOM_FLD]
        # /!\ FeatureClassToNumPyArray can't handle <null> /!\
        data = arcpy.da.FeatureClassToNumPyArray(
            self.source,
            flds,
            u'{depcom_fld} in ({depcoms})'.format(
                depcoms=self.depcoms.strg,
                depcom_fld=DEPCOM_FLD),
            skip_nulls = False,
            null_value = 0)
        # insert data in template xlsx
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(self.path)
        ws_data = wb.Worksheets('data')
        ws_data.Range('A2:D'+str(data.size+1)).Value = data
        ws_data.Range('E2:G2').AutoFill(
            Destination=ws_data.Range('E2:G'+str(data.size+1)))
        # update titles
        ws_titles = wb.Worksheets('titres')
        ws_titles.Range('A1').Value = self.title
        # update all pivot tables
        for sh in wb.Worksheets:
            for pt in sh.PivotTables():
                pt.RefreshTable()
        #   - export pdf and close xlsx (no save required)
        wb.ExportAsFixedFormat(
            From=1, To=4, Type=0, Filename=self.pdf, IgnorePrintAreas=False)
        wb.Close(True)
        # flag self as processed
        self.processed = True

        
class Report(object):
    """
        Builds Report object
    """

    # class variables (?)
    # ----------------
    # - spot datasource (in case)
    # /!\ wont work properly /!\:
    # data sources are already set in templates
    # they must be adjusted manually before anything

    source = MOS_PATH

    #   - templates root folder
    tpl_root = (
        ur'J:\Etudes\laufma\Python26\site-packages\mezcal\templates\pandemic_edition')

    #   - templates
    templates = [
        ur'st_10_tpl.mxd',
        ur'st_20_tpl.mxd',
        ur'evo_10_20_tpl.mxd',
        ur'data_tpl.xlsx']      # <- xlsx template

    def __init__(self, depcoms_path):
        """
            initiates a new Report object
            :depcoms_path: path to depcoms file
            
            templates are defined as class variables
            
            The Report object contains all params
            to build the real report but
            no processing is done at this point.

            :param depcoms_path: path to depcoms
            :type depcoms_path: unicode
            :return: report object, will be processed if process() is triggered
            :rtype: Report instance
        """
        # initiate depcoms
        self.depcoms = Depcoms(depcoms_path)
        # sets paths
        self.title = naming(depcoms_path)
        # pages created after templates list
        self.allpages = [
            Page(
                self.source, self.depcoms,
                os.path.join(self.tpl_root, template))
            for template in self.templates]

    def process(self, *pages):
        """
            method to effectively build the report
            :param pages: pages to process. If None: all pages are processed
            :type pages: basestrings

            Available pages: 'st_00', 'st_05', 'st_10', 'st_15', 'evo_00_10',
            'evo_05_15', 'evo_00_05', 'evo_05_10', 'evo_10_15', 'ortho_05',
            'ortho_16', 'data'
            
            usage:
            ------
            will process given pages
            >>> rep.process()
            will process all pages
            >>> rep.process('evo_10_15', 'data')
        """
        # default behaviour: process all pages
        if not pages:
            pages = (
                'st_10', 'st_20', 'evo_10_20', 'data')
        # process all pages specified in pages param
        for page in self.allpages:
            # filter by pages
            if page.name in pages:
                print u'\tprocessing {0}...'.format(page.name),
                page.process_page()
                print u'done.'
        # assemble all in one big pdf
        root, _ = os.path.split(self.depcoms.path)
        self.pdf = os.path.join(root, 'pdf', self.title + '.pdf')
        pdfs = [page.pdf for page in self.allpages if page.processed]
        pdfDoc = arcpy.mapping.PDFDocumentCreate(self.pdf)
        for page in pdfs:
            pdfDoc.appendPages(page)
        pdfDoc.saveAndClose()
        print u'\treport {0} done.'.format(self.title)
        del pdfDoc

def main(depcoms):
    """
        starts the whole damn thing
        :param depcoms: path to depcoms file
        :type depcoms: unicode
    """
    if not isinstance(depcoms, unicode):
        depcoms = depcoms.decode('utf-8')
    myReport = Report(depcoms)
    myReport.process()
    

if __name__ == "__main__":

    # retrieve script paramater
    depcoms = sys.argv[1]
    # run main
    main(depcoms)
