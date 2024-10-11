from docx import Document
from pprint import pprint


class Metadata:

    def __init__ (self, filename):
        self._filename = filename
        self._structures = list()
        self.get_data()

    def read_docx_metadata(self):
        doc = Document(self._filename)
        mt = dict()
        for attr in self._structures:
           mt[attr] = getattr(doc.core_properties, attr)
        return mt
    
    def get_data(self): 
        doc = Document(self._filename)
        for attr in dir(doc.core_properties):
            if not callable(attr) and not str(attr).startswith('__') and not str(attr).startswith('_'):
                self._structures.append(attr)

    def write_docx_metadata(self, data):
        doc = Document(self._filename)
        for attr in data:
            setattr(doc.core_properties, attr, data[attr])
        doc.save(self._filename)
        
    
if __name__ == "__main__":
    metadata = Metadata('src/data/file.docx')
    pprint(metadata.read_docx_metadata())
    metadata.write_docx_metadata({'title': 'algo', 'author': 'author'})