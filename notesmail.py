import datetime
from win32com.client import DispatchEx
from win32com.client import makepy

makepy.GenerateFromTypeLibSpec('Lotus Domino Objects')
makepy.GenerateFromTypeLibSpec('Lotus Notes Automation Classes')

class NotesMail():
    """
     发送读取邮件有关的操作
    """
    def __init__(self, server, file):
        """初始化连接
            @param server
             服务器名
            @param file
             数据文件名
        """
        self.session = DispatchEx('Notes.NotesSession')
        self.server = self.session.GetEnvironmentString("MailServer", True)
        self.db = self.session.GetDatabase(server, file)
        self.db.OPENMAIL
        self.myviews = []

    def send_mail(self, receiver, subject, body=None):
        """发送邮件
            @param receiver: 收件人
            @param subject: 主题
            @param body: 内容
        """
        doc = self.db.CREATEDOCUMENT
        doc.sendto = receiver
        doc.Subject = subject
        if body:
            doc.Body = body
        doc.SEND(0, receiver)

    def get_views(self):
        for view in self.db.Views:
            if view.IsFolder:
                self.myviews.append(view.name)
    def make_document_generator(self, view_name):
        self.__get_folder()
        folder = self.db.GetView(view_name)
        if not folder:
            raise Exception('Folder {0} not found. '.format(view_name))
        document = folder.GetFirstDocument
        while document:
            yield document
            document = folder.GetNextDocument(document)

    def read_mail(self):
        for document in self.make_document_generator('Mine'):
            result = self.extract_documet(document)

        print(result)
        """
        self.d
        dirc = self.session.GetDbDirectory("Domino/罗皓")
        d = self.db.OpenMailDatabase
        view = d.GetView("$inbox")
        doc = view.getfirstdocument()
        """
    def extract_documet(self, document):
        """提取Document
        """
        result = {}
        result['subject'] = document.GetItemValue('Subject')[0].strip()
        result['date'] = document.GetItemValue('PostedDate')[0]
        result['From'] = document.GetItemValue('From')[0].strip()
        result['To'] = document.GetItemValue('SendTo')
        result['body'] = document.GetItemValue('Body')[0].strip()

        return result

def main():
    mail = NotesMail("ZH45XXSS05/BOC", 'mail\lh7167.nsf')
    #mail = NotesMail('C:\\Program Files (x86)\\IBM\\Lotus\\Notes\\Data\\as_罗皓.nsf')
    #mail.read_mail()
    mail.send_mail('603750199@qq.com', 'Good afternoon', 'Wish you a good day')

if __name__ == '__main__':
    main()

