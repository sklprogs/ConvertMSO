#!/usr/bin/python3
# -*- coding: UTF-8 -*-

import os
from skl_shared.localize import _
import skl_shared.shared as sh


class File:
    
    def __init__(self):
        self.source = ''
        self.target = ''
        self.source_ext = ''
        self.target_ext = ''
        self.folder = ''
        self.Failed = False
        self.Skipped = False
        self.source_size = 0
        self.target_size = 0



class Convert:
    
    def __init__(self,folder,Debug=False):
        self.set_values()
        self.folder = folder
        self.Debug = Debug
    
    def set_values(self):
        self.timer = sh.Timer('[ConvertMSO] controller.Convert')
        self.Success = True
        self.Debug = False
        self.folder = ''
        self.ifiles = []
        self.allowed_ext = ('.doc','.xls')
        self.target_ext = ('.odt','.ods')
    
    def delete(self):
        f = '[ConvertMSO] controller.Convert.delete'
        if not self.Success:
            sh.com.cancel(f)
            return
        count = 0
        for ifile in self.ifiles:
            ''' 'target', 'Skipped' and 'Failed' are being checked
                when setting 'target_size'.
            '''
            if ifile.target_size > 0:
                if sh.File(ifile.source).delete():
                    count += 1
        sh.com.rep_matches(f,count)
    
    def check_output(self):
        f = '[ConvertMSO] controller.Convert.check_output'
        if not self.Success:
            sh.com.cancel(f)
            return
        for ifile in self.ifiles:
            if ifile.target and not os.path.exists(ifile.target):
                ifile.Failed = True
    
    def _convert(self,ifile):
        f = '[ConvertMSO] controller.Convert._convert'
        format_ = ifile.target_ext.replace('.','')
        ''' - The target is silently overwritten by LibreOffice
            - If '--outdir' is not given, the *script* folder will be
              used!
        '''
        args = ['--headless','--convert-to',format_
               ,'--outdir',ifile.folder
               ]
        return sh.Launch(ifile.source,True).launch_app('libreoffice',args)
    
    def convert(self):
        f = '[ConvertMSO] controller.Convert.convert'
        if not self.Success:
            sh.com.cancel(f)
            return
        for ifile in self.ifiles:
            if ifile.target and not ifile.Skipped:
                if self._convert(ifile):
                    ifile.target_size = sh.File(ifile.target).get_size()
                else:
                    ifile.Failed = True
    
    def _get_processed(self):
        count = 0
        for ifile in self.ifiles:
            if ifile.target and not ifile.Skipped and not ifile.Failed:
                count += 1
        return count
    
    def _get_skipped(self):
        count = 0
        for ifile in self.ifiles:
            if ifile.Skipped:
                count += 1
        return count
    
    def _get_failed(self):
        count = 0
        for ifile in self.ifiles:
            if ifile.Failed:
                count += 1
        return count
    
    def _get_souce_size(self):
        size = 0
        for ifile in self.ifiles:
            if ifile.target and not ifile.Skipped and not ifile.Failed:
                size += ifile.source_size
        return size
    
    def _get_target_size(self):
        size = 0
        for ifile in self.ifiles:
            if ifile.target and not ifile.Skipped and not ifile.Failed:
                size += ifile.target_size
        return size
    
    def report(self):
        f = '[ConvertMSO] controller.Convert.report'
        if not self.Success:
            sh.com.cancel(f)
            return
        mes = []
        processed = [ifile for ifile in self.ifiles if ifile.target]
        failed = [ifile for ifile in self.ifiles if ifile.Failed]
        skipped = [ifile for ifile in self.ifiles if ifile.Skipped]
        sub = _('Files in total: {}').format(len(self.ifiles))
        mes.append(sub)
        sub = _('Processed files: {}').format(self._get_processed())
        mes.append(sub)
        sub = _('Skipped files: {}').format(self._get_skipped())
        mes.append(sub)
        sub = _('Failed files: {}').format(self._get_failed())
        mes.append(sub)
        old_size = self._get_souce_size()
        new_size = self._get_target_size()
        size_diff = old_size - new_size
        # Avoid ZeroDivisionError
        if old_size:
            percent = round((100*size_diff)/old_size)
        else:
            percent = 0
        old_size = sh.com.get_human_size(old_size,True)
        new_size = sh.com.get_human_size(new_size,True)
        size_diff = sh.com.get_human_size(size_diff,True)
        sub = _('Processed data: {}').format(old_size)
        mes.append(sub)
        sub = _('Converted data: {}').format(new_size)
        mes.append(sub)
        sub = _('Compression: {} ({}%)').format(size_diff,percent)
        mes.append(sub)
        delta = sh.com.get_human_time(self.timer.end())
        sub = _('The operation has taken {}').format(delta)
        mes.append(sub)
        mes = '\n'.join(mes)
        sh.objs.get_mes(f,mes).show_info()
    
    def check(self):
        self.idir = sh.Directory(self.folder)
        self.Success = self.idir.Success
    
    def run(self):
        self.timer.start()
        self.check()
        self.set_files()
        self.set_target()
        self.convert()
        self.check_output()
        self.delete()
        self.debug()
        self.report()
    
    def debug(self):
        f = '[ConvertMSO] controller.Convert.debug'
        if not self.Success:
            sh.com.cancel(f)
            return
        if not self.Debug:
            sh.com.rep_lazy(f)
            return
        source_exts = []
        target_exts = []
        source_sizes = []
        target_sizes = []
        sources = []
        targets = []
        failed = []
        skipped = []
        nos = [i + 1 for i in range(len(self.ifiles))]
        for ifile in self.ifiles:
            source_exts.append(ifile.source_ext)
            target_exts.append(ifile.target_ext)
            sources.append(ifile.source)
            targets.append(ifile.target)
            failed.append(ifile.Failed)
            skipped.append(ifile.Skipped)
            source_sizes.append(ifile.source_size)
            target_sizes.append(ifile.target_size)
        headers = (_('#'),'EXT1',_('SOURCE'),_('SIZE{}').format(1)
                  ,'EXT2',_('TARGET'),_('SIZE{}').format(2)
                  ,_('SKIPPED'),_('FAILED')
                  )
        iterable = [nos,source_exts,sources,source_sizes,target_exts
                   ,targets,target_sizes,skipped,failed
                   ]
        mes = sh.FastTable (headers = headers
                           ,iterable = iterable
                           ,maxrow = 30
                           ,maxrows = 1000
                           ,FromEnd = True
                           ).run()
        sh.com.run_fast_debug(f,mes)
    
    def _get_target_ext(self,ext):
        # The calling code already ensures input presence
        index_ = self.allowed_ext.index(ext)
        return self.target_ext[index_]
    
    def set_target(self):
        f = '[ConvertMSO] controller.Convert.set_target'
        if not self.Success:
            sh.com.cancel(f)
            return
        for ifile in self.ifiles:
            if ifile.source_ext in self.allowed_ext:
                ifile.source_size = sh.File(ifile.source).get_size()
                ipath = sh.Path(ifile.source)
                target_ext = self._get_target_ext(ifile.source_ext)
                basename = ipath.get_filename() + target_ext
                ifile.target_ext = target_ext
                ifile.folder = ipath.get_dirname()
                ifile.target = os.path.join(ifile.folder,basename)
                if os.path.exists(ifile.target):
                    ifile.Skipped = True
    
    def set_files(self):
        f = '[ConvertMSO] controller.Convert.set_files'
        if not self.Success:
            sh.com.cancel(f)
            return
        files = sh.Directory(self.folder).get_subfiles()
        if files:
            for file in files:
                ifile = File()
                ifile.source = file
                ifile.source_ext = sh.Path(file).get_ext_low()
                self.ifiles.append(ifile)
        else:
            self.Success = False
            sh.com.rep_empty(f)


if __name__ == '__main__':
    f = '[ConvertMSO] controller.__main__'
    sh.com.start()
    folder = sh.Home().get_home()
    Convert(folder,False).run()
    sh.com.end()
