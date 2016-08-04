# -*- coding: utf-8 -*-
import time
import struct
import socket
import hashlib
import base64
import sys
from select import select
import re
import logging
from threading import Thread
import signal
import json
from win32com.client import Dispatch
from collections import OrderedDict
import urllib2
import os
import pythoncom
# Constants
MAGIC = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"
TEXT = 0x01
BINARY = 0x02

logging.basicConfig(filename="C:\\HSS_SOCKET.log",level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# WebSocket implementation
class WebSocket(object):
    handshake = (
        "HTTP/1.1 101 Web Socket Protocol Handshake\r\n"
        "Upgrade: WebSocket\r\n"
        "Connection: Upgrade\r\n"
        "Sec-WebSocket-Accept: %(acceptstring)s\r\n"
        "Server: TestTest\r\n"
        "Access-Control-Allow-Origin: http://localhost\r\n"
        "Access-Control-Allow-Credentials: true\r\n"
        "\r\n"
    )

    # Constructor
    def __init__(self, client, server):
        self.client = client
        self.server = server
        self.handshaken = False
        self.header = ""
        self.data = ""
        self.message = ""
        # handle COM object in multithread
        pythoncom.CoInitialize()
        self.shell = Dispatch("Shell.Application")
        self.wscript = Dispatch("WScript.Shell")

    # Serve this client
    def feed(self, data):

        # If we haven't handshaken yet
        if not self.handshaken:
            logging.debug("No handshake yet")
            self.header += data
            if self.header.find('\r\n\r\n') != -1:
                parts = self.header.split('\r\n\r\n', 1)
                self.header = parts[0]
                if self.dohandshake(self.header, parts[1]):
                    logging.info("Handshake successful")
                    self.handshaken = True

        # We have handshaken
        else:
            logging.debug("Handshake is complete")            
            # Decode the data that we received according to section 5 of RFC6455            
            recv = self.decodeCharArray(data)
        
            try:
                recv_data = json.loads(''.join(recv).strip(), object_pairs_hook=OrderedDict)
                if recv_data:
                    if recv_data['action'] == 'PUSH_F1':
                        self.setIEElementByName(recv_data['content'])
                    if recv_data['action'] == 'PUSH_DOCUMENT':
                        path = self.saveTaskDocument(recv_data)
                        if path:
                            command = r"%d"+path+r"{ENTER}"
                            self.selectFile(command, recv_data['doc']['filename'])
                            self.message = "Lưu document thành công !"                            
            except Exception, e:
                logging.error(str(e))
                self.message = str(e)
                pass
            # Send our reply
            self.sendMessage(self.message)
        
    def saveTaskDocument(self, data):        
        doc = data['doc']
        upload_windows = None
        for win in self.shell.Windows():
            if win.Name == "Windows Internet Explorer" and win.LocationURL == "http://cfris02.fecredit.com.vn/VPBank/adddoc.jsp":
                upload_windows = win                
                break
        if not upload_windows:
            self.message = "Chưa mở form upload chứng từ Omnidocs"
            return False        
        # Check if appid is matches
        omni_appid = win.Document.getElementsByTagName("td")[2].childnodes[0].innerHTML
        if omni_appid.replace("&nbsp; ","") != doc['app_id']:
            self.message = "App ID không đúng, mời bạn kiểm tra lại!"
            return False
        # Check Documet desc
        desc = win.Document.getElementsByTagName("td")[6].childnodes[0] .innerHTML
        if desc.replace("&nbsp; ","") != doc['desc']:
            self.message = "Sai loại chứng từ !"
            return False
        # Disguise as browser
        request_headers = {
            "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Encoding":"gzip, deflate, sdch",
            "Accept-Language":"en-US,en;q=0.8,vi;q=0.6",
            "Cache-Control":"max-age=0",
            "Connection":"keep-alive",
            "Cookie":data['cookie'],
            "DNT":"1",            
            "Upgrade-Insecure-Requests":"1",
            "User-Agent":"Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.106 Safari/537.36"
        }
        fileurl = doc['fileurl']
        path = r"C:\HSS\{0}\{1}".format(doc['customer_id'], doc['app_id'])
        try:
            request = urllib2.Request(fileurl, headers=request_headers)
            f = urllib2.urlopen(request).read()
            if not os.path.exists(path):
                os.makedirs(path)
            with open(r"{0}\{1}".format(path, doc['filename']),"wb") as local_file:
                local_file.write(f)
            return path
        except urllib2.HTTPError, e:
            logging.error(e.headers)
            logging.error(e.code + e.msg)
            self.message = e.code + e.msg
            return False        

    # Select file to upload
    def selectFile(self, command, filename):
        browser = {"IE": 0}
        def _clickBrowseInput(args):            
            pythoncom.CoInitialize()
            shell = Dispatch("Shell.Application")            
            for win in self.shell.Windows():
                if win.Name == "Windows Internet Explorer" and win.LocationURL == "http://cfris02.fecredit.com.vn/VPBank/adddoc.jsp":
                    args['IE'] = win                
                    break
            if args['IE']:
                args['IE'].Document.getElementById("inputfile").click()
                        
        # Start click "Browse" button at new thread because this action requires select file to complete
        clickBrowseInput = Thread(target = _clickBrowseInput, args = (browser, ))           
        clickBrowseInput.start()
        # Choose file to upload
        time.sleep(1)        
        if browser['IE']:
            pythoncom.CoInitialize()
            wscript = Dispatch("WScript.Shell")           
            while not wscript.AppActivate("Choose File to Upload"):        
                time.sleep(1)                    
            wscript.SendKeys(command, 1)
            
            time.sleep(1)
            wscript.SendKeys("%n", 1)
            logging.info("Choosing file %s" % filename)
            wscript.SendKeys(filename, 1)
            wscript.SendKeys("{ENTER}", 1)
            #time.sleep(1)        
            while wscript.AppActivate("Choose File to Upload"):               
                time.sleep(1)
            currentIE = self.shell.Windows().Item(self.shell.Windows().Count - 1)
            currentIE.Document.getElementsByTagName("a")[1].click()
            self.message = "Gửi document tới OmniDocs thành công !"
        else:
            self.message = "Không tìm thấy form upload Omnidocs !"
        # End Thread
        clickBrowseInput.join()  
            
    
    def setIEElementByName(self, data):                
        values = data['values']        
        title = data['title']
        self.message = u"Không tìm thấy trang %s".encode("utf-8") % title.encode("utf-8")
        for win in self.shell.Windows():            
            if win.Name == "Windows Internet Explorer":                
                if win.Document.getElementsByTagName("title")[0].innerHTML == title:
                    self.message = u"Gửi thông tin %s thành công".encode("utf-8") % title.encode("utf-8")
                    for fieldName in values:                        
                        el = win.Document.getElementsByName(fieldName)
                        if el.length:
                            #Change field's value
                            if 'val' in values[fieldName] and values[fieldName]['val']:
                                el[0].value = values[fieldName]['val']                                
                            #Fire event of field
                            if 'event' in values[fieldName] and values[fieldName]['event']:
                                el[0].FireEvent(values[fieldName]['event'])
                                if values[fieldName]['event'] == 'onclick':
                                    while win.ReadyState != 4:
                                        time.sleep(0.5)
                                if values[fieldName]['event'] == 'onchange':
                                    #Wait for field many2one finish loading
                                    time.sleep(1)
                                    popup = self.shell.Windows().Item(self.shell.Windows().Count-1)
                                    while popup.ReadyState !=4:
                                        time.sleep(0.5)
                                    if popup.Document.Title != title:
                                        while popup in self.shell.Windows():
                                            time.sleep(0.5)  
        return True

    # Stolen from http://www.cs.rpi.edu/~goldsd/docs/spring2012-csci4220/websocket-py.txt
    def sendMessage(self, s):
        """
        Encode and send a WebSocket message
        """

        # Empty message to start with
        message = ""

        # always send an entire message as one frame (fin)
        b1 = 0x80

        # in Python 2, strs are bytes and unicodes are strings
        if type(s) == unicode:
            b1 |= TEXT
            payload = s.encode("UTF8")

        elif type(s) == str:
            b1 |= TEXT
            payload = s

        # Append 'FIN' flag to the message
        message += chr(b1)

        # never mask frames from the server to the client
        b2 = 0

        # How long is our payload?
        length = len(payload)
        if length < 126:
            b2 |= length
            message += chr(b2)

        elif length < (2 ** 16) - 1:
            b2 |= 126
            message += chr(b2)
            l = struct.pack(">H", length)
            message += l

        else:
            l = struct.pack(">Q", length)
            b2 |= 127
            message += chr(b2)
            message += l

        # Append payload to message
        message += payload

        # Send to the client
        self.client.send(str(message))

    # Stolen from http://stackoverflow.com/questions/8125507/how-can-i-send-and-receive-websocket-messages-on-the-server-side
    def decodeCharArray(self, stringStreamIn):

        # Turn string values into opererable numeric byte values
        byteArray = [ord(character) for character in stringStreamIn]
        datalength = byteArray[1] & 127
        indexFirstMask = 2

        if datalength == 126:
            indexFirstMask = 4
        elif datalength == 127:
            indexFirstMask = 10

        # Extract masks
        masks = [m for m in byteArray[indexFirstMask: indexFirstMask + 4]]
        indexFirstDataByte = indexFirstMask + 4

        # List of decoded characters
        decodedChars = []
        i = indexFirstDataByte
        j = 0

        # Loop through each byte that was received
        while i < len(byteArray):
            # Unmask this byte and add to the decoded buffer
            decodedChars.append(chr(byteArray[i] ^ masks[j % 4]))
            i += 1
            j += 1

        # Return the decoded string
        return decodedChars

    # Handshake with this client
    def dohandshake(self, header, key=None):

        logging.debug("Begin handshake: %s" % header)

        # Get the handshake template
        handshake = self.handshake

        # Step through each header
        for line in header.split('\r\n')[1:]:
            name, value = line.split(': ', 1)

            # If this is the key
            if name.lower() == "sec-websocket-key":
                # Append the standard GUID and get digest
                combined = value + MAGIC
                response = base64.b64encode(hashlib.sha1(combined).digest())

                # Replace the placeholder in the handshake response
                handshake = handshake % {'acceptstring': response}

        logging.debug("Sending handshake %s" % handshake)
        self.client.send(handshake)
        return True

    def onmessage(self, data):
        # logging.info("Got message: %s" % data)
        self.send(data)

    def send(self, data):
        logging.info("Sent message: %s" % data)
        self.client.send("\x00%s\xff" % data)    
    
    def close(self):
        self.client.close()


# WebSocket server implementation
class WebSocketServer(object):
    # Constructor
    def __init__(self, bind, port, cls):
        self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        #self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        self.socket.bind((bind, port))
        self.bind = bind
        self.port = port
        self.cls = cls
        self.connections = {}
        self.listeners = [self.socket]

    # Listen for requests
    def listen(self, backlog=5):        
        self.socket.listen(backlog)
        logging.info("Listening on %s" % self.port)

        # Keep serving requests
        self.running = True
        while self.running:
    
            # Find clients that need servicing
            rList, wList, xList = select(self.listeners, [], self.listeners, 1)
            for ready in rList:
                if ready == self.socket:
                    logging.debug("New client connection")
                    client, address = self.socket.accept()
                    fileno = client.fileno()
                    self.listeners.append(fileno)
                    self.connections[fileno] = self.cls(client, self)
                else:
                    logging.debug("Client ready for reading %s" % ready)
                    client = self.connections[ready].client
                    try:
                        data = client.recv(1024)
                    except Exception, e:
                        logging.error(str(e))
                        data = ''
                    fileno = client.fileno()
                    if data:
                        self.connections[fileno].feed(data)
                    else:
                        logging.debug("Closing client %s" % ready)
                        self.connections[fileno].close()
                        del self.connections[fileno]
                        self.listeners.remove(ready)

            # Step though and delete broken connections
            for failed in xList:
                if failed == self.socket:
                    logging.error("Socket broke")
                    for fileno, conn in self.connections:
                        conn.close()
                    self.running = False
                    
import win32serviceutil
import win32service
import win32event
import servicemanager


class HSS_SocketService (win32serviceutil.ServiceFramework):
    _svc_name_ = "HSS_SERVICE"
    _svc_display_name_ = "HSS Socket Service"
    
    def __init__(self,args):
        win32serviceutil.ServiceFramework.__init__(self,args)
        self.stop_event = win32event.CreateEvent(None,0,0,None)
        self.stop_requested = False

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
        logging.info('Stopping service ...')
        self.stop_requested = True

    def SvcDoRun(self):
        #servicemanager.LogMsg(
        #    servicemanager.EVENTLOG_INFORMATION_TYPE,
        #    servicemanager.PYS_SERVICE_STARTED,
        #    (self._svc_name_,'')
        #)
        self.main()

    def main(self):                
        server = WebSocketServer("", 8888, WebSocket)        
        server_thread = Thread(target=server.listen, args=[5])
        server_thread.start()	
        
        while True:
            time.sleep(3)            
            if self.stop_requested:
                break         
        return

if __name__ == '__main__':    
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(HSS_SocketService)
        servicemanager.StartServiceCtrlDispatcher()
    else:        
        win32serviceutil.HandleCommandLine(HSS_SocketService)
