# -*- coding: utf-8 -*-
"""
@Time      : 2020/12/16 15:19
@Author    : William.sv@icloud.com
@File      : docxParser.py
@ Software : PyCharm
@Desc      : 获取word文档内容及图片
"""

import xml.dom.minidom
import zipfile
import os
import sys



class DocxParser:

    def get_content(self,file_path):
        if not os.path.exists(file_path):
            print('文件不存在，请检查后重试~')
            sys.exit()
        if '.docx' not in file_path:
            print('仅能处理docx后缀的文件，请检查后重试')
            sys.exit()
        article_content = self.__read_resource(file_path)
        print(article_content)
        return article_content

    def __read_resource(self,file_path):
        content = ''
        rels = {}
        with zipfile.ZipFile(file_path, 'r') as zip_file:
            file_list = zip_file.namelist()
            document_xml_name = ''  # 段落文件
            document_xml_rels = ''  # media文件
            for i in file_list:
                if 'document.xml' in i :
                    document_xml_name = i
                if 'document.xml.rels' in i :
                    document_xml_rels = i
            with zip_file.open(document_xml_name, mode='r') as resource_document:
                DOMTree = xml.dom.minidom.parseString(resource_document.read())
                collection = DOMTree.documentElement
                wp = collection.getElementsByTagName('w:p')
                for item in wp:
                    section = ''
                    p = ''
                    picture_id = ''
                    for i in item.childNodes:
                        wt = i.getElementsByTagName('w:t')
                        pic = i.getElementsByTagName('pic:blipFill')
                        if len(wt) == 0 and len(pic) == 0:
                            continue
                        if len(wt) > 0:
                            p = p + wt[0].childNodes[0].data  # 获取段落
                        if len(pic) > 0:
                            picture_id = picture_id + ' {' + pic[0].getElementsByTagName('a:blip')[0].attributes[
                                'r:embed'].value + '} '  # 获取图片资源位
                    if p != '':
                        section = '<p>' + p + '</p>'
                    if len(picture_id) > 0:
                        section = section + picture_id
                    content = content + section
            with zip_file.open(document_xml_rels, mode='r') as resource_rels:
                DOMTree = xml.dom.minidom.parseString(resource_rels.read())
                collection = DOMTree.documentElement
                wp = collection.getElementsByTagName('Relationship')
                for i in wp:
                    id = i.attributes['Id'].value
                    target = i.attributes['Target'].value
                    if 'media/' in target:
                        rels[id] = target.split('/')[-1].split('.')[0]
        for k,v in rels.items():
            if '{' + k + '}' in content:
                content = content.replace('{' + k + '}','{'+ v +'}')
        return content

    def __parser_content(self,element):
        DOMTree = xml.dom.minidom.parseString(element)
        collection = DOMTree.documentElement
        content = ''
        wp = collection.getElementsByTagName('w:p')
        for item in wp:
            section = ''
            p = ''
            picture_id = ''
            for i in item.childNodes:
                wt = i.getElementsByTagName('w:t')
                pic = i.getElementsByTagName('pic:blipFill')
                if len(wt) == 0 and len(pic) == 0:
                    continue
                if len(wt) > 0:
                    p = p + wt[0].childNodes[0].data  # 获取段落
                if len(pic) > 0:
                    picture_id = picture_id + ' {' + pic[0].getElementsByTagName('a:blip')[0].attributes['r:embed'].value + '} '  # 获取图片资源位
            if p != '':
                section = '<p>' + p + '</p>'
            if len(picture_id) > 0:
                section = section + picture_id
            content = content + section

        return content

    def __parse_rels(self,element):
        DOMTree = xml.dom.minidom.parse(element)
        collection = DOMTree.documentElement
        wp = collection.getElementsByTagName('Relationship')
        rels = {}
        for i in wp:
            id = i.attributes['Id'].value
            target = i.attributes['Target'].value
            if 'media' in target:
                rels[id] = target
        return rels

    def get_media(self,file_path,save_path):
        with zipfile.ZipFile(file_path) as zip_file:
            file_list = zip_file.namelist()
            media = []
            for i in file_list:
                if 'word' in i and 'media' in i:
                    media.append(i)
            if media != '':
                for i in media:
                    zip_file.extract(i, save_path)
                print('文件已保存在' + os.path.join(save_path,'word','media') + '中！')