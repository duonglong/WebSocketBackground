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
import win32api
from collections import OrderedDict
import urllib2
import os
import pythoncom
# Constants
MAGIC = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"
TEXT = 0x01
BINARY = 0x02

#logging.basicConfig(filename="C:\\HSS_SOCKET.log",level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")

# ====================================================WebSocket implementation================================================

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
        self.timeout = 30
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
                            self.selectFile(path, recv_data['doc']['filename'], recv_data['doc']['app_id'])
                            self.message = "Lưu document thành công !"                            
            except Exception, e:
                logging.exception(str(e))
                self.message = str(e)
                pass
            # Send our reply
            self.sendMessage(self.message)
        
    def saveTaskDocument(self, data):
        logging.info("Begin downloading file")
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
        path = r"C:\hanel\{0}\{1}".format(doc['customer_id'], doc['app_id'])
        logging.info("Saving file to %s" % path)
        try:
            request = urllib2.Request(fileurl, headers=request_headers)
            f = urllib2.urlopen(request).read()
            if not os.path.exists(path):
                os.makedirs(path)
            with open(r"{0}\{1}".format(path, doc['filename']),"wb") as local_file:
                local_file.write(f)
            return path
        except urllib2.HTTPError, e:
            logging.exception(e.headers)
            logging.exception(e.code + e.msg)
            self.message = e.code + e.msg
            return False        

    # Select file to upload    
    def selectFile(self, path, filename, app_id):
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
        time.sleep(1.5)
        if browser['IE']:
            pythoncom.CoInitialize()            
            wscript = Dispatch("WScript.Shell")
            start = 0
            while not wscript.AppActivate("Choose File to Upload"):
                if start == self.timeout:
                    logging.debug("Time out when waiting for upload popup !")
                    self.message = "Time out when waiting for upload popup !"
                    return
                time.sleep(1)
                start += 1
                                
            wscript.SendKeys("%d", 0)
            win32api.Sleep(500)            
            wscript.SendKeys(path, 0)
            win32api.Sleep(500)            
            wscript.SendKeys("{ENTER}", 0)
            win32api.Sleep(500)            
            wscript.SendKeys("%n", 0)
            win32api.Sleep(500)       
            wscript.SendKeys(filename, 0)
            win32api.Sleep(500)            
            wscript.SendKeys("{ENTER}", 0)
            win32api.Sleep(500)

            #Wait till done selecting file
            while wscript.AppActivate("Choose File to Upload"):               
                time.sleep(1)
            for win in self.shell.Windows():
                if win.Name == "Windows Internet Explorer" and win.Document.Title == "VPBank@%s" % app_id:
                    win.Document.getElementsByTagName("a")[1].click()
            self.message = "Gửi document tới OmniDocs thành công !"
        else:
            self.message = "Không tìm thấy form upload Omnidocs !"
        # End Thread
        clickBrowseInput.join()  
            
    
    def setIEElementByName(self, data):                
        values = data['values']        
        title = data['title']
        self.message = u"Không tìm thấy trang %s".encode("utf-8") % title.encode("utf-8")
        IE = None
        for win in self.shell.Windows():
            print win.Name
            if win.Name == "Windows Internet Explorer" and win.Document.Title == title:                
                IE = win
        if IE:
            for fieldName in values:                     
                el = IE.Document.getElementsByName(fieldName)
                if el.length:
                    #Change field's value
                    if 'val' in values[fieldName] and values[fieldName]['val']:
                        logging.debug("Filling field %s" % fieldName)
                        el[0].value = values[fieldName]['val']                                
                    #Fire event of field
                    if 'event' in values[fieldName] and values[fieldName]['event']:                        
                        el[0].FireEvent(values[fieldName]['event'])
                        if values[fieldName]['event'] == 'onclick':
                            logging.debug("Fire event %s on field %s" % ("onclick",fieldName))
                            start = 0
                            while IE.ReadyState != 4:
                                time.sleep(0.5)
                                start += 0.5
                                if start == self.timeout:
                                    logging.debug("Time out when waiting for page ready !")
                                    self.message = "Time out when waiting for page ready !"
                                    return False
                        if values[fieldName]['event'] == 'onchange':
                            logging.debug("Fire event %s on field %s" % ("onchange",fieldName))
                            #Wait for field many2one finish loading                            
                            popup = self.shell.Windows().Item(self.shell.Windows().Count - 1)
                            start = 0
                            while popup in self.shell.Windows() and popup.ReadyState !=4:
                                time.sleep(0.5)
                                start += 0.5
                                if start == self.timeout:
                                    logging.debug("Time out when waiting for popup ready !")
                                    self.message = "Time out when waiting for popup ready !"
                                    return False
                            if popup in self.shell.Windows() and popup.Document.Title != title:
                                while popup in self.shell.Windows():
                                    time.sleep(0.5)
            self.message = u"Gửi thông tin %s thành công".encode("utf-8") % title.encode("utf-8")
        else:
            self.message = u"Chưa mở tab %s".encode("utf-8") % title.encode("utf-8")
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
                        logging.exception(str(e))
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
                    logging.exception("Socket broke")
                    for fileno, conn in self.connections:
                        conn.close()
                    self.running = False       
# ============================================================= SYSTRAY ================================================================
# Credit : Simon Brunning - simon@brunningonline.net

         
import win32con
import win32gui_struct
try:
    import winxpgui as win32gui
except ImportError:
    import win32gui

class SysTrayIcon(object):
    '''TODO'''
    QUIT = 'QUIT'
    SPECIAL_ACTIONS = [QUIT]
    
    FIRST_ID = 1023
    
    def __init__(self,
                 icon,
                 hover_text,
                 menu_options,
                 on_quit=None,
                 default_menu_index=None,
                 window_class_name=None,):
        
        self.icon = icon
        self.hover_text = hover_text
        self.on_quit = on_quit
        
        menu_options = menu_options + (('Quit', None, self.QUIT),)
        self._next_action_id = self.FIRST_ID
        self.menu_actions_by_id = set()
        self.menu_options = self._add_ids_to_menu_options(list(menu_options))
        self.menu_actions_by_id = dict(self.menu_actions_by_id)
        del self._next_action_id
        
        
        self.default_menu_index = (default_menu_index or 0)
        self.window_class_name = window_class_name or "SysTrayIconPy"
        
        message_map = {win32gui.RegisterWindowMessage("TaskbarCreated"): self.restart,
                       win32con.WM_DESTROY: self.destroy,
                       win32con.WM_COMMAND: self.command,
                       win32con.WM_USER+20 : self.notify,}
        # Register the Window class.
        window_class = win32gui.WNDCLASS()
        hinst = window_class.hInstance = win32gui.GetModuleHandle(None)
        window_class.lpszClassName = self.window_class_name
        window_class.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW;
        window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        window_class.hbrBackground = win32con.COLOR_WINDOW
        window_class.lpfnWndProc = message_map # could also specify a wndproc.
        classAtom = win32gui.RegisterClass(window_class)
        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(classAtom,
                                          self.window_class_name,
                                          style,
                                          0,
                                          0,
                                          win32con.CW_USEDEFAULT,
                                          win32con.CW_USEDEFAULT,
                                          0,
                                          0,
                                          hinst,
                                          None)
        win32gui.UpdateWindow(self.hwnd)
        self.notify_id = None
        self.refresh_icon()
        
        win32gui.PumpMessages()

    def _add_ids_to_menu_options(self, menu_options):
        result = []
        for menu_option in menu_options:
            option_text, option_icon, option_action = menu_option
            if callable(option_action) or option_action in self.SPECIAL_ACTIONS:
                self.menu_actions_by_id.add((self._next_action_id, option_action))
                result.append(menu_option + (self._next_action_id,))
            elif non_string_iterable(option_action):
                result.append((option_text,
                               option_icon,
                               self._add_ids_to_menu_options(option_action),
                               self._next_action_id))
            else:
                print 'Unknown item', option_text, option_icon, option_action
            self._next_action_id += 1
        return result
        
    def refresh_icon(self):
        # Try and find a custom icon
        hinst = win32gui.GetModuleHandle(None)
        if os.path.isfile(self.icon):
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            hicon = win32gui.LoadImage(hinst,
                                       self.icon,
                                       win32con.IMAGE_ICON,
                                       0,
                                       0,
                                       icon_flags)
        else:
            print "Can't find icon file - using default."
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

        if self.notify_id: message = win32gui.NIM_MODIFY
        else: message = win32gui.NIM_ADD
        self.notify_id = (self.hwnd,
                          0,
                          win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                          win32con.WM_USER+20,
                          hicon,
                          self.hover_text)
        win32gui.Shell_NotifyIcon(message, self.notify_id)

    def restart(self, hwnd, msg, wparam, lparam):
        self.refresh_icon()

    def destroy(self, hwnd, msg, wparam, lparam):
        if self.on_quit: self.on_quit(self)
        nid = (self.hwnd, 0)
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostQuitMessage(0) # Terminate the app.

    def notify(self, hwnd, msg, wparam, lparam):
        if lparam==win32con.WM_LBUTTONDBLCLK:
            self.execute_menu_option(self.default_menu_index + self.FIRST_ID)
        elif lparam==win32con.WM_RBUTTONUP:
            self.show_menu()
        elif lparam==win32con.WM_LBUTTONUP:
            pass
        return True
        
    def show_menu(self):
        menu = win32gui.CreatePopupMenu()
        self.create_menu(menu, self.menu_options)
        #win32gui.SetMenuDefaultItem(menu, 1000, 0)
        
        pos = win32gui.GetCursorPos()
        # See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/menus_0hdi.asp
        win32gui.SetForegroundWindow(self.hwnd)
        win32gui.TrackPopupMenu(menu,
                                win32con.TPM_LEFTALIGN,
                                pos[0],
                                pos[1],
                                0,
                                self.hwnd,
                                None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
    
    def create_menu(self, menu, menu_options):
        for option_text, option_icon, option_action, option_id in menu_options[::-1]:
            if option_icon:
                option_icon = self.prep_menu_icon(option_icon)
            
            if option_id in self.menu_actions_by_id:                
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hbmpItem=option_icon,
                                                                wID=option_id)
                win32gui.InsertMenuItem(menu, 0, 1, item)
            else:
                submenu = win32gui.CreatePopupMenu()
                self.create_menu(submenu, option_action)
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hbmpItem=option_icon,
                                                                hSubMenu=submenu)
                win32gui.InsertMenuItem(menu, 0, 1, item)

    def prep_menu_icon(self, icon):
        # First load the icon.
        ico_x = win32api.GetSystemMetrics(win32con.SM_CXSMICON)
        ico_y = win32api.GetSystemMetrics(win32con.SM_CYSMICON)
        hicon = win32gui.LoadImage(0, icon, win32con.IMAGE_ICON, ico_x, ico_y, win32con.LR_LOADFROMFILE)

        hdcBitmap = win32gui.CreateCompatibleDC(0)
        hdcScreen = win32gui.GetDC(0)
        hbm = win32gui.CreateCompatibleBitmap(hdcScreen, ico_x, ico_y)
        hbmOld = win32gui.SelectObject(hdcBitmap, hbm)
        # Fill the background.
        brush = win32gui.GetSysColorBrush(win32con.COLOR_MENU)
        win32gui.FillRect(hdcBitmap, (0, 0, 16, 16), brush)
        # unclear if brush needs to be feed.  Best clue I can find is:
        # "GetSysColorBrush returns a cached brush instead of allocating a new
        # one." - implies no DeleteObject
        # draw the icon
        win32gui.DrawIconEx(hdcBitmap, 0, 0, hicon, ico_x, ico_y, 0, 0, win32con.DI_NORMAL)
        win32gui.SelectObject(hdcBitmap, hbmOld)
        win32gui.DeleteDC(hdcBitmap)
        
        return hbm

    def command(self, hwnd, msg, wparam, lparam):
        id = win32gui.LOWORD(wparam)
        self.execute_menu_option(id)
        
    def execute_menu_option(self, id):
        menu_action = self.menu_actions_by_id[id]      
        if menu_action == self.QUIT:
            win32gui.DestroyWindow(self.hwnd)
        else:
            menu_action(self)
            
def non_string_iterable(obj):
    try:
        iter(obj)
    except TypeError:
        return False
    else:
        return not isinstance(obj, basestring)


if __name__ == '__main__':
    import itertools, glob
    
    icons = itertools.cycle(glob.glob('*.ico'))
    hover_text = "HSS WebSocket"

    obj = {"server": WebSocketServer("", 8888, WebSocket)}
    server_thread = Thread(target=obj["server"].listen, args=[5])
    server_thread.start()
    win32api.MessageBox(0, 'HSS WEBSOCKET OPENED', 'INFO',0x00001000)
    #def runSocket(sysTrayIcon):
    #    win32api.MessageBox(0, 'Socket is running', 'WARNING')
    #    server_thread = Thread(target=obj["server"].listen, args=[5])
    #    server_thread.start()
    #       while True:
    #       time.sleep(100)
        
    #def stopSocket(sysTrayIcon):
    #    logging.info("Shutting down...")
    #    obj["server"].running = False
    #    win32api.MessageBox(0, 'Closing HSS Socket', 'WARNING',0x00001000)
    #    sys.exit()
        
    menu_options = (
                    #('Close Socket', icons.next(), stopSocket),                    
                   )
    def bye(sysTrayIcon):
        logging.info("Shutting down...")
        win32api.MessageBox(0, 'HSS WEBSOCKET CLOSED', 'WARNING',0x00001000)
        obj["server"].running = False
        sys.exit()
    
    SysTrayIcon(icons.next(), hover_text, menu_options, on_quit=bye, default_menu_index=1)  
