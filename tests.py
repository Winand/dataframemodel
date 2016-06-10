# -*- coding: utf-8 -*-
"""
Created on Wed Jun  8 08:46:24 2016

@author: МакаровАС
"""

import pythoncom
import win32com.client as w32client
import win32ui
import win32file
import win32pipe
import pywintypes
import zlib
import os
import time

wagons = """56739915
52289378
53825006
58036534
54902358
53825634
52479508
54902044
57444143
61932778"""


def dispatch_running(disp_names):
    "Dispatch object in ROT"
    if not isinstance(disp_names, (list, tuple)):
        disp_names = (disp_names,)
    context = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()
    for i in rot:
        dp = i.GetDisplayName(context, None)
        if dp in disp_names:
            obj = rot.GetObject(i).QueryInterface(pythoncom.IID_IDispatch)
            print("Found running object", dp)
            return w32client.Dispatch(obj)


class NamedPipe():
    lineterm, prefix = '\r\n', r'\\.\pipe\%s'

    def __init__(self, name, encoding='cp1251'):
        self.name, self.encoding = name, encoding
        self.pipe = self.create_pipe(name)

    def path(self):
        return self.prefix % self.name
#        pipename = r"\\.\pipe"+win32file.GetFileInformationByHandleEx(
#                                            self.pipe, win32file.FileNameInfo)

    def read_message(self, prefix="~start~", postfix="~end~"):
        "Read pipe from `prefix` (optionally) to `postfix` (required)"
        (_start_, _end_), data = map(lambda s: s.encode(self.encoding),
                                     (prefix, postfix+self.lineterm)), b''
        while not data.endswith(_end_):
            data += self.read()
        data = data[len(_start_)*data.startswith(_start_):]
        return data[:-len(_end_)].decode(self.encoding)

    def read(self):
        win32pipe.ConnectNamedPipe(self.pipe, None)
        data = b''
        while True:
            try:
                data += win32file.ReadFile(self.pipe, 4096)[1]
            except pywintypes.error as e:
                if e.winerror == 109:  # Error Broken Pipe
                    break
                else:
                    raise e
        win32pipe.DisconnectNamedPipe(self.pipe)
        return data

    def create_pipe(self, name):
        return win32pipe.CreateNamedPipe(
            self.prefix % name, win32pipe.PIPE_ACCESS_DUPLEX,
            win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_WAIT, 1, 65536, 65536,
            300, None)

    def __del__(self):
        self.close()

    def close(self):
        win32file.CloseHandle(self.pipe)

    def write(self, text, close=False):
        "close - close pipe to emulate EOF"
        win32pipe.ConnectNamedPipe(self.pipe, None)
        text += self.lineterm
        win32file.WriteFile(self.pipe, text.encode(self.encoding))
        if close:
            # http://comp.os.linux.questions.narkive.com/2AW9g5yn/sending-an-eof-to-a-named-pipe
            self.close()
        else:
            win32pipe.DisconnectNamedPipe(self.pipe)


class WaitForModification():
    "Wait for /path/ file to be modified before exiting /with/ block"
    def __init__(self, path):
        self.path = path
        self.mtime = not os.path.exists(path) or os.path.getmtime(path)

    def __enter__(self):
        pass

    def __exit__(self, *args):
        while True:
            try:
                if os.path.getmtime(self.path) != self.mtime:
                    break
            except:
                pass
            time.sleep(0.01)


class SAS():
    NAMEDPIPE = "PySAS"
    SAS_DISP_NAMES = ("!{89FA3E2A-43F9-43E4-B1A2-DAC2CC90B89C}",  # 9.4
                      "!{C3B368D4-C09D-49D8-A7A6-12FACEFA6F38}",  # 9.3
                      "!{87CE93EC-4802-49EA-B8C9-F7A4F41612BB}",
                      "!{375AA476-B299-462B-A9EB-A817C9DC1EA8}",
                      "!{CF6FE983-0701-4B8C-B044-3CFE12977E48}",
                      "!{61149040-2E51-11CF-B3F0-444553540000}",  # 8.2
                      "!{E4EF3900-BA1B-11CF-9C0F-0020AF34BE48}",
                      "!{A18FB580-317C-11D0-A908-0020AFCEF803}",)
    CONN_VAR = "CONND"  # SAS variable used to check the connection status

    def __init__(self, show=False):
        self.sas = sas = dispatch_running(self.SAS_DISP_NAMES)
        if show:
            sas.Visible = True
        # There's no Wait property in 9.4 example, don't rely on it
        # http://support.sas.com/documentation/cdl/en/hostwin/67962/HTML/default/viewer.htm#p14xcdjd86154dn1kcdtorc69uwu.htm
        # https://v8doc.sas.com/sashtml/win/zeedback.htm
        sas.Wait = True
        # Setup named pipe communication
        self.pipe = NamedPipe(self.NAMEDPIPE)
        self.submit("filename CONN namepipe '%s' client retry=-1;" %
                    self.pipe.path())
#        self.submit("*); */; /*’*/ /*”*/; %mend;")
        self.workdir = self.submit(ret_val="_wd=%sysfunc(pathname(work))")

    def submit(self, text="", ret_val=None, remote=False, wait=True):
        """
        ret_val (name[=val]) - define local value, submit, get value
        remote - rsubmit statements
        wait - block until SAS is ready
        """
        assert text or ret_val
        if ret_val:
            macro_var = [i.strip() for i in ret_val.split("=")]
            if len(macro_var) == 2:  # set macro variable default value
                text = "%let {}={}; ".format(*macro_var) + text
            elif len(macro_var) > 2:
                raise ValueError("Error, ret_val format is name[=val]")
        if remote:  # wrap statements in rsubmit block
            text = "rsubmit; %s endrsubmit;" % text
        try:
            self.sas.Submit(text)
            if wait:
                self.wait()
        except:
            print("Connection to SAS client failed")
        if ret_val:
            return self.get_sas_var(macro_var[0])

    def wait(self):
        while self.sas.Busy:
            win32ui.PumpWaitingMessages(0, -1)

    def get_sas_var(self, name):
        "Read SAS variable via named pipe"
        # RECFM=N doesn't work w/ pipes, so \r\n is added to the end anyway
        self.submit("data _null_;"
                    "file CONN;"
                    "put '&%s'@;"
                    "put '~end~';"
                    "run;" % name, wait=False)
        ret = self.pipe.read_message()
        self.wait()
        return ret

    def include(self, text, remote=False):
        "Submit statements using %INCLUDE"
        pipe = NamedPipe("inc_pipe")
        self.submit("filename inc_pipe namepipe '%s' client;" % pipe.path())
        self.submit("%include inc_pipe /SOURCE2;", wait=False)
        if remote:
            text = "rsubmit; %s endrsubmit;" % text
        pipe.write(text, close=True)
        self.wait()

    def set_debug(self, debug):
        self.debug = debug
        printto = "filename abyss dummy; proc printto {%d}; run;" % debug
        self.submit(printto.format("log=abyss", ""))
        options = "options {%d}source; options {%d}notes;" % (debug, debug)
        self.submit(options.format("no", ""))
        if self.is_connected():
            self.submit(options.format("no", ""), remote=True)

    def is_connected(self):
        login_status = self.submit("signon macvar=%s;" % self.CONN_VAR,
                                   ret_val=self.CONN_VAR+"=1")
        return int(login_status) == 2  # Already connected

    def signon(self, server, creds):
        us, sv = creds.popitem(), server.popitem()
        server_str = "%s %d" % sv
        # Generate 8 symbol name
        comp_name = hex(zlib.crc32(server_str.encode()) >> 4)[1:]
        if not all(sum((us, sv), ())):
            print("Error:Check server address and credentials")
            return
        self.submit(r"%let {}={};".format(comp_name, server_str))
        singon_stmt = ("signon connectremote={} user={} "
                       "password={} macvar=%s noscript;" % self.CONN_VAR)
        login_status = self.submit(singon_stmt.format(comp_name, *us),
                                   ret_val=self.CONN_VAR+"=1")
        return int(login_status) in (0, 2)


sas = SAS(show=True)
sas.set_debug(True)
print('connect...')
r = sas.signon({"10.200.8.1": 5050}, {"DDWHRZ9E": "msi947"})
if r:
    print("select...")
    q = """
%macro rs(num);
proc sql;
connect to remote(server=X1242ACB dbms=db2);
   execute(CALL @LRAO04.SBDSP021(&num)) by remote;
   create table temp_rs as select * from connection to remote(select DATE_REG from SESSION.SBDSP021);
disconnect from remote;
quit;
%mend;
"""
    q = """
%macro rs(num);
%syslput num=&num;
rsubmit; %rs_r; endrsubmit;
%mend;
rsubmit;
%macro rs_r;
proc sql;
connect to db2;
   execute(CALL @LRAO04.SBDSP021(&num)) by db2;
   create table temp_rs as select * from connection to db2(select * from SESSION.SBDSP021);
disconnect from db2;
quit;
proc download data=temp_rs out=temp_rs status=no; run;
%mend;
endrsubmit;
"""
    sas.include(q)
    for i in wagons.splitlines():
        a1 = time.clock()
        with WaitForModification(sas.workdir+r"\temp_rs.sas7bdat"):
            sas.submit("%rs({});".format(i))
        with open(sas.workdir+r"\temp_rs.sas7bdat", 'rb') as f:
            print(f.read()[:5])
        print(i, time.clock() - a1)
#        time.sleep(0.05)
print('ok')
