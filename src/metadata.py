from docx import Document
from pptx import Presentation
from openpyxl import load_workbook
from pprint import pprint


class Metadata:

    def __init__ (self, filename):
        self._filename = filename
        self._structures = list()
        self._get_data()

    def read_docx_metadata(self):
        """Reads metadata from a docx file"""
        doc = Document(self._filename)
        mt = dict()
        for attr in self._structures:
           mt[attr] = getattr(doc.core_properties, attr)
        return mt
    
    def write_docx_metadata(self, data):
        """Writes metadata to a docx file"""
        doc = Document(self._filename)
        for attr in data:
            setattr(doc.core_properties, attr, data[attr])
        doc.save(self._filename)

    def _get_data(self): 
        """Get the structures of the metadata"""

        file_extension = self._filename.split('.')[-1]
        if file_extension in ['docx', 'docm', 'dotx', 'dotm']: 
            file= Document(self._filename)
            for attr in dir(file.core_properties):
                if not callable(attr) and not str(attr).startswith('__') and not str(attr).startswith('_'):
                    self._structures.append(attr)
            
        elif file_extension  in ['pptx', 'pptm', 'potx', 'potm', 'ppsx', 'ppsm']: 
            file = Presentation(self._filename)
            for attr in dir(file.core_properties):
                if not callable(attr) and not str(attr).startswith('__') and not str(attr).startswith('_'):
                    self._structures.append(attr)
        elif file_extension in ['xlsx', 'xlsm', 'xltx', 'xltm']:
            file = load_workbook(self._filename)

        
            for attr in dir(file.properties):
                if not callable(attr) and not str(attr).startswith('__') and not str(attr).startswith('_'):
                    self._structures.append(attr)



    def read_pptx_metadata(self):
        """Reads metadata from a pptx file"""
        ppt = Presentation(self._filename)
        mt = dict()
        for attr in self._structures:
           mt[attr] = getattr(ppt.core_properties, attr)
        return mt
        
    def write_pptx_metadata(self, data):
        """Writes metadata to a pptx file"""
        ppt = Presentation(self._filename)
        for attr in data:
            setattr(ppt.core_properties, attr, data[attr])
        ppt.save(self._filename)

    def read_xlsx_metadata(self):
        """Reads metadata from a xlsx file"""
        wb = load_workbook(self._filename)
        mt = dict()
        for attr in self._structures:
           mt[attr] = getattr(wb.properties, attr)
        return mt
    
    def write_xlsx_metadata(self, data):
        """Writes metadata to a xlsx file"""
        wb = load_workbook(self._filename)
        for attr in data:
            setattr(wb.properties, attr, data[attr])
        wb.save(self._filename)

if __name__ == "__main__":
    metadata = Metadata('src/data/file.docx')
    pprint(metadata.read_docx_metadata())
    metadata.write_docx_metadata({'title': 'algo', 'author': 'author'})