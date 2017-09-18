from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
from django.utils.encoding import smart_str
import sys,os,time
import win32com.client,pythoncom
import threading
import random

def hello(request):
    return HttpResponse("Hello world ! ")
shalfile = ""
lock = threading.Lock()
def upload_file(request):
    if request.method == 'POST' and request.FILES['myfile']:
        global shalfile
        myfile = request.FILES['myfile']
        shalfile = request.POST['shalfile']
        print time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        lock.acquire()
        result = mindle_convert(myfile)
        lock.release()
        print time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())      

    return HttpResponse(False)

def testupload_file(request):
    if request.method == 'POST' and request.FILES['myfile']:
        global shalfile,lock
        myfile = request.FILES['myfile']
        #shalfile = request.POST['shalfile']
        shalfile = str(time.time())
        print shalfile
        print time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
      
        lock.acquire()
        result = mindle_convert(myfile)
        lock.release()
        print time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())      
        return HttpResponse(result)
        # if flag == 0:
        #     lock = threading.Lock()
        #     t = ThreadConvert(queue,lock)
        #     t.setDaemon(True)
        #     t.start()
        #     flag = flag+1
        #     queue.put(myfile,True,None)
        #     queue.join()
        # else:
        #     flag = flag+1
        #     queue.put(myfile,True,None)
        #     queue.join()
       

    return HttpResponse(False)

def mindle_convert(myfile):

    basePath = "D:\\ConvertDocuments\\ConvertDocuments\\static\\ConvertFiles\\"
    if os.path.isfile(basePath+ myfile.name) == False:
        fs = FileSystemStorage(location=basePath)
        filename = fs.save(myfile.name, myfile)
        #time.sleep(1)
        uploaded_file_url = fs.url(filename)           
    if os.path.isfile(basePath+myfile.name+".pdf") == False:           
        try:
            ext = os.path.splitext(basePath+ myfile.name)[1]   
            global shalfile
            ret = ConvertDocument(basePath+myfile.name,shalfile+str(random.random())+".pdf", ext.lower())
            if ret:
                return HttpResponse()
        except BaseException:
            return HttpResponse(False)
    else:
            return HttpResponse()

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
    basePath = "D:\\ConvertDocuments\\ConvertDocuments\\static\\ConvertFiles\\"
    pythoncom.CoInitialize()  
    if ext == ".docx" or ext == ".doc" or ext == ".wps" :      
        try:
            wdFormatPDF = 17 
            word = win32com.client.Dispatch("Word.Application") 
            word.Visible = 0 
            word.DisplayAlerts = 0
            try:
                doc = word.Documents.Open(in_file)      
            except IOError as error:
                 print (error)
                 doc.Close()
                 word.Quit()                   
            doc.SaveAs(basePath+out_file,FileFormat=wdFormatPDF)
            doc.Close()
            word.Quit()
        except pythoncom.com_error as error:
            print (error)
            print (vars(error))
            print (error.args)
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
        ppt.SaveAs(basePath+out_file,FileFormat=32)
        ppt.Close()
        powerpoint.Quit()
    return True
# print time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) 

def threadtest(request,test):
    return HttpResponse(False)

