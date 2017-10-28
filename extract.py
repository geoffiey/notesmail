import os
import tempfile
from exceptions import *

class Extract():
    def __init__(self, document):
        """初始化
            @param document
             The Notes Document that will be extract
        """
        self.document = document

    def __get_temp_path(self):
        temp_index, temp_path = tempfile.mkstemp()
        os.close(temp_index)
        return temp_path

    def get_attachment(self, filepath=None):
        """Get attachments of a document
            @param filepath
             The path want to save file
             if not given, then current directory will be used
        """
        attachment_packs = []

        for item in range(len(self.document.Items)):
            t_item = self.document.Items[item]
            if t_item.Name == '$FILE':
                attachment_path = self.__get_temp_path()
                filename = t_item.Values[0]
                filebase, separator, file_extension = filename.rpartition('.')
                attahment = self.document.GetAttachment(filename)
                attahment.ExtractFile(attachment_path)
                attachment_content = open(attachment_path, 'rb').read()
                os.remove(attachment_path)
                attachment_packs.append((filebase, file_extension, attachment_content))

        if len(attachment_packs) == 0:
            raise NoAttachmentException('No attachment in document. ')

        for pack in attachment_packs:
            filename = pack[0] + '.' + pack[1]
            if filepath:
                filename = filepath + filename
            with open(filename, 'wb') as fh:
                fh.write(pack[2])

    def extract(self):
        """提取Document
            @param document
             notes文档
            @return dict.
             subject -> 主题
             date ->日期
             From -> 发件人
             To -> 收件人
             body -> 主体
        """
        result = {}
        result['subject'] = self.document.GetItemValue('Subject')[0].strip()
        result['date'] = self.document.GetItemValue('PostedDate')[0]
        result['From'] = self.document.GetItemValue('From')[0].strip()
        result['To'] = self.document.GetItemValue('SendTo')
        result['body'] = self.document.GetItemValue('Body')[0].strip()

        return result
