import datetime
from win32com.client import DispatchEx
from win32com.client import makepy

from extract import Extract

makepy.GenerateFromTypeLibSpec('Lotus Domino Objects')
makepy.GenerateFromTypeLibSpec('Lotus Notes Automation Classes')

class NotesMail():
    """
     发送读取邮件有关的操作
    """
    def __init__(self, server, file):
        """Initialize
            @param server
             Server's name of Notes
            @param file
             Your data file, usually ends with '.nsf'
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

    def get_documents(self, view_name):
        """
            @return generator
        """
        documents = []
        folder = self.db.GetView(view_name)
        if not folder:
            raise Exception('Folder {0} not found. '.format(view_name))
        document = folder.GetFirstDocument
        while document:
            documents.append(document)
            document = folder.GetNextDocument(document)

        return documents

    def read_mail(self, view, attachment=False):
        """Read the latest mail
            @param view
             The view(fold) to access
            @param attachment
             Boolean, whether get attachment
            @return, dict
             Info of a mail
        """
        result = {}

        documents = self.get_documents(view)
        latest_document = documents[-1:][0]
        extra_obj = Extract(latest_document)
        result = extra_obj.extract()
        if attachment:
            extra_obj.get_attachment()

        return result

def main():
    pass
if __name__ == '__main__':
    main()

