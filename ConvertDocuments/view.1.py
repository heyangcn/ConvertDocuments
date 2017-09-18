from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
from django.utils.encoding import smart_str
import sys,os,time
import win32com.client,pythoncom
import Queue
import threading

queue = Queue.Queue()
flag = 0
def hello(request):
    return HttpResponse("Hello world ! ")


def upload_file(request):
    global queue
    global flag
    if request.method == 'POST' and request.FILES['myfile']:
        myfile = request.FILES['myfile']
      
        if flag == 0:
            lock = threading.Lock()
            t = ThreadConvert(queue,lock)
            t.setDaemon(True)
            t.start()
            flag = flag+1
            queue.put(myfile,True,None)
            queue.join()
        else:
            flag = flag+1
            queue.put(myfile,True,None)
            queue.join()
       

    return HttpResponse(123)


def mindle_convert(myfile):

    basePath = "D:/ConvertDocuments/ConvertDocuments/static/ConvertFiles/"
    if os.path.isfile(basePath+ myfile.name) == False:
        fs = FileSystemStorage(location=basePath)
        filename = fs.save(myfile.name, myfile)
        uploaded_file_url = fs.url(filename)           
    if os.path.isfile(basePath+myfile.name+".pdf") == False:           
        try:
            ext = os.path.splitext(basePath+ myfile.name)[1]   
            ret = ConvertDocument(basePath+myfile.name, myfile.name+".pdf", ext.lower())
            if ret:
                return HttpResponse(True)
        except BaseException:
            return HttpResponse(BaseException)
    else:
            return HttpResponse(True)

def  returnPdfFile(filename):
     with open('D:/ConvertDocuments/ConvertDocuments/static/ConvertFiles/'+filename, 'rb') as pdf:
         response = HttpResponse(pdf.read(),content_type='application/pdf')
         response['Content-Disposition'] = 'attachment;filename='+filename
         return response

def  show_pdffile(request):
    filename = "7930145e4637063cb740c765c59222ab9a35edd4.docx.pdf"
    with open('D:/ConvertDocuments/ConvertDocuments/static/ConvertFiles/'+filename, 'rb') as pdf:
        response = HttpResponse(pdf.read(),content_type='application/pdf')
        response['Content-Disposition'] = 'inline;filename='+filename
        return response



        
def ConvertDocument(in_file, out_file, ext):
    # print time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    # in_file = os.path.abspath(sys.argv[1])
    # out_file = os.path.abspath(sys.argv[2])
    # ext = sys.argv[3]
    # return os.path.abspath(os.curdir)
    basePath = "D:/ConvertDocuments/ConvertDocuments/static/ConvertFiles/"
    pythoncom.CoInitialize()
    if ext == ".docx" or ext == ".doc" or ext == ".wps" :
        wdFormatPDF = 17 
        word = win32com.client.Dispatch("Word.Application") 
        doc = word.Documents.Open(in_file)
        doc.SaveAs(basePath+out_file,wdFormatPDF)
        doc.Close()
        word.Quit()
    elif ext == ".xlsx" or ext == ".xls" :
        xlApp = win32com.client.Dispatch("Excel.Application")
        books = xlApp.Workbooks.Open(in_file)
        ws = books.Worksheets[0]
        ws.Visible = 1
        ws.ExportAsFixedFormat(0,basePath+out_file)
        books.Close()
        # xlApp.Quit()
    elif ext == ".ppt" or ext == ".pptx" :
        pptFormatPDF = 32
        powerpoint = win32com.client.Dispatch("Powerpoint.Application") 
        ppt = powerpoint.Presentations.Open(in_file)
        powerpoint.Visible = 1
        ppt.SaveAs(basePath+out_file,32)
        ppt.Close()
        powerpoint.Quit()
    return True
# print time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) 

def threadtest(request,test):
    return HttpResponse(False)




class ThreadConvert(threading.Thread):
  """Threaded Url Grab"""
  def __init__(self, queue,lock):
    threading.Thread.__init__(self)
    self.queue = queue
    self.lock = lock

  def run(self):
    global flag
    while True:
      #grabs item from queue     
      self.lock.acquire()
      print time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) +"  begintime"
      myfile = self.queue.get()
      mindle_convert(myfile)
      print time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) +"  endtime"
      self.lock.release()
      #signals to queue job is done
      time.sleep(1)
      self.queue.task_done()
      flag = flag-1
