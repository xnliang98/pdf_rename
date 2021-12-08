#!/usr/bin/env python3

"""
A script to batch rename PDF files based on metadata/XMP title and author
Requirements:
    - PDFMiner: https://github.com/pdfminer/pdfminer.six
    - xmp: lightweight XMP parser from
        http://blog.matt-swain.com/post/25650072381/
            a-lightweight-xmp-parser-for-extracting-pdf-metadata-in
"""


NAME = 'pdf-title-rename'
VERSION = '0.1.0'
DATE = '2017-07-15'


import os
import sys
import argparse
import subprocess
import glob

# PDF and metadata libraries
from pdfminer.pdfparser import PDFParser, PDFSyntaxError
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1

"""
Parses XMP metadata from PDF files.
By Matt Swain. Released under the MIT license:
http://blog.matt-swain.com/post/25650072381/a-lightweight-xmp-parser-for-extracting-pdf-metadata-in
"""

from collections import defaultdict
from xml.etree import ElementTree as ET

RDF_NS = '{http://www.w3.org/1999/02/22-rdf-syntax-ns#}'
XML_NS = '{http://www.w3.org/XML/1998/namespace}'
NS_MAP = {
    'http://www.w3.org/1999/02/22-rdf-syntax-ns#'    : 'rdf',
    'http://purl.org/dc/elements/1.1/'               : 'dc',
    'http://ns.adobe.com/xap/1.0/'                   : 'xap',
    'http://ns.adobe.com/pdf/1.3/'                   : 'pdf',
    'http://ns.adobe.com/xap/1.0/mm/'                : 'xapmm',
    'http://ns.adobe.com/pdfx/1.3/'                  : 'pdfx',
    'http://prismstandard.org/namespaces/basic/2.0/' : 'prism',
    'http://crossref.org/crossmark/1.0/'             : 'crossmark',
    'http://ns.adobe.com/xap/1.0/rights/'            : 'rights',
    'http://www.w3.org/XML/1998/namespace'           : 'xml'
}

class XmpParser(object):
    """
    Parses an XMP string into a dictionary.
    Usage:
        parser = XmpParser(xmpstring)
        meta = parser.meta
    """

    def __init__(self, xmp):
        self.tree = ET.XML(xmp)
        self.rdftree = self.tree.find(RDF_NS+'RDF')

    @property
    def meta(self):
        """ A dictionary of all the parsed metadata. """
        meta = defaultdict(dict)
        for desc in self.rdftree.findall(RDF_NS+'Description'):
            for el in list(desc):
                ns, tag =  self._parse_tag(el)
                value = self._parse_value(el)
                meta[ns][tag] = value
        return dict(meta)

    def _parse_tag(self, el):
        """ Extract the namespace and tag from an element. """
        ns = None
        tag = el.tag
        if tag[0] == "{":
            ns, tag = tag[1:].split('}',1)
            if ns in NS_MAP:
                ns = NS_MAP[ns]
        return ns, tag

    def _parse_value(self, el):
        """ Extract the metadata value from an element. """
        if el.find(RDF_NS+'Bag') is not None:
            value = []
            for li in el.findall(RDF_NS+'Bag/'+RDF_NS+'li'):
                value.append(li.text)
        elif el.find(RDF_NS+'Seq') is not None:
            value = []
            for li in el.findall(RDF_NS+'Seq/'+RDF_NS+'li'):
                value.append(li.text)
        elif el.find(RDF_NS+'Alt') is not None:
            value = {}
            for li in el.findall(RDF_NS+'Alt/'+RDF_NS+'li'):
                value[li.get(XML_NS+'lang')] = li.text
        else:
            value = el.text
        return value

def xmp_to_dict(xmp):
    """ Shorthand function for parsing an XMP string into a python dictionary. """
    return XmpParser(xmp).meta


def _sanitize(s):
    keep = [" ", ".", "_", "-", "\u2014"]
    return "".join([x for x in s if x in keep or x.isalnum()]).strip()

def _get_metadata(f):
    parser = PDFParser(f)
    try:
        doc = PDFDocument(parser)
    except PDFSyntaxError:
        return {}
    parser.set_document(doc)

    if not hasattr(doc, 'info') or len(doc.info) == 0:
        return {}
    return doc

def _resolve_objref(ref):
    if hasattr(ref, 'resolve'):
        return ref.resolve()
    return ref

def _au_last_name(name):
    return name.split()[-1]

def _get_xmp_metadata(doc):
    t = a = None
    metadata = resolve1(doc.catalog['Metadata']).get_data()
    try:
        md = xmp_to_dict(metadata)
    except:
        return t, a

    try:
        t = md['dc']['title']['x-default']
    except TypeError:
        # The 'title' field might be a string or bytes instead of a dict
        # https://github.com/jdmonaco/pdf-title-rename/issues/7
        titleval = md['dc']['title']
        if type(titleval) is str:
            t = titleval
        elif type(titleval) is bytes:
            t = titleval.decode()
    except KeyError:
        pass

    try:
        a = md['dc']['creator']
    except KeyError:
        pass
    else:
        if type(a) is bytes:
            a = a.decode('utf-8')
        if type(a) is str:
            a = [a]
        a = list(filter(bool, a))  # remove None, empty strings, ...
        if len(a) > 1:
            a = '%s %s' % (_au_last_name(a[0]), _au_last_name(a[-1]))
        elif len(a) == 1:
            a = _au_last_name(a[0])
        else:
            a = None

    return t, a

def _get_info(fn):
    title = author = None
    with open(fn, "rb") as f:
        doc = _get_metadata(f)
        
        info = doc.info[0]
        if 'Title' in info:
            ti = _resolve_objref(info['Title'])
            try:
                title = ti.decode('utf-8')
            except AttributeError:
                pass
            except UnicodeDecodeError:
                print(' -- Could not decode title bytes: %r' % ti)
        
        if 'Author' in info:
            au = _resolve_objref(info['Author'])
            try:
                author = au.decode('utf-8')
            except AttributeError:
                pass
            except UnicodeDecodeError:
                print(' -- Could not decode title bytes: %r' % au)
                
        if 'Metadata' in doc.catalog:
            xmpt, xmpa = _get_xmp_metadata(doc)
            xmpt = _resolve_objref(xmpt)
            xmpa = _resolve_objref(xmpa)
            if xmpt:
                title = xmpt
            if xmpa:
                author = xmpa
                
    if type(title) is str:
        title = title.strip()
        if title.lower() == 'untitled':
            title = None

    # if .interactive:
    #     title, author = ._interactive_info_query(fn, title, author)

    return title, author

def _new_filename(title, author=None):
    title = title.lower().replace(":", "")
    
    n = _sanitize(title)
    n = "-".join(title.strip().split(" "))
    if author is not None:
        author = author.lower()
        n = '%s.%s' % (_sanitize(author).lower(), n)
    n = '%s.pdf' % n[:250]  # limit filenames to ~255 chars
    return n

def main(dir, is_author=False, destination=None):
    
    files = glob.glob(os.path.join(dir, f"*.pdf"))
    Ntot = len(files)
    Nmissing = 0
    Nfiled = 0
    Nerrors = 0
    Nrenamed = 0
    # print(files)
    for f in files:
        print(f)
        root, ext = os.path.splitext(f)
        path, base = os.path.split(root)
        title, author = _get_info(f)
        # print(title, author)
        if author and not title:
            title = base
        if not (author or title):
            print(' -- Could not find metadata in the file')
            Nmissing += 1
            continue
        if is_author:
            newf = os.path.join(path, _new_filename(title, author))
        else:
            newf = os.path.join(path, _new_filename(title))
        # print(newf)
        try:
            os.rename(f, newf)
        except OSError:
            print(' -- Error renaming file, maybe it moved?')
            Nerrors += 1
            continue
        else:
            Nrenamed += 1

        if destination:
            if subprocess.call(['mv', newf, destination]) == 0:
                print(' -- Filed to', destination)
                Nfiled += 1
            else:
                print(' -- Error moving file')
                Nerrors += 1
            print(' - Filed: %d' % Nfiled)
    print('Processed %d files:' % Ntot)
    print(' - Renamed: %d' % Nrenamed)
    print(' - Missing metadata: %d' % Nmissing)
    print(' - Errors: %d' % Nerrors)
        
    return None

dir = sys.argv[1] 
main(dir)