
# coding: utf-8

# In[ ]:

import argparse, json, os, sys
import gspread
import httplib2
from oauth2client.client import OAuth2WebServerFlow
from oauth2client import file, tools
from myauth import *

 
class Unbuffered(object):
   def __init__(self, stream):
       self.stream = stream
   def write(self, data):
       self.stream.write(data)
       self.stream.flush()
   def __getattr__(self, attr):
       return getattr(self.stream, attr)


class MyCreds (object): 
    def __init__ (self, access_token=None): 
        self.access_token = access_token 


def blanksheet(asheet):
    import re
    apat = re.compile(r'.*')
    blanks = asheet.findall(apat)
    for blankme in blanks:
        blankme.value = ''
    try:
        asheet.update_cells(blanks)
    except:
        return False
    else:
        return True


def updatecells(credentials, cells, asheet):
    if credentials.invalid:
        gs.refreshcreds(credentials)
    try:
        asheet.update_cells(cells)
    except Exception as e:
        import traceback
        print(traceback.print_exc())


def refreshcreds(credentials):
    http = credentials.authorize(http = httplib2.Http())
    credentials.refresh(http)


def localtime():
    from datetime import datetime
    from dateutil.tz import tzutc, tzlocal
    utc = datetime.now(tzutc())
    local = utc.astimezone(tzlocal()).strftime("%Y-%m-%d %H:%M:%S")
    return(local)


def connection_test(sheet='GoodSheet', table='Sheet1', rows=360, secs=10):
    # Insert a timestamp every minute to the end of a Google Sheet.
    from time import sleep
    gssdoc = connect.open(sheet) # Check for bad filename
    asheet = gssdoc.worksheet(table) # Check for bad tablename
    if not blanksheet(asheet):
        raise SystemExit('Could not blank spreadsheet')
    for i in range(rows):
        print('%s, ' % i, end='')
        if credentials.invalid:
            refreshcreds(credentials) # Don't let OAuth2 login expire
        try:
            asheet.update_cell(i+1, 1, localtime())
        except Exception as e:
            import traceback
            print(traceback.print_exc())
            break
        sleep(secs)
    print("Done")


def popucolumn(asheet, rows=1000, secs=0):
    arange = 'A1:A%s' % rows
    cells = asheet.range(arange)
    for cell in cells:
        col = cell.col
        row = cell.row
        if col == 1:
            cell.value = row
    try:
        updatecells(credentials, cells, asheet)
        return True
    except:
        return False


def copycolumn(asheet, rows=1000, secs=0):
    arange = 'A1:B%s' % rows
    cells = asheet.range(arange)
    for cell in cells:
        col = cell.col
        row = cell.row
        if col == 2:
            cell.value = val
        val = cell.value
    try:
        updatecells(credentials, cells, asheet)
        return True
    except:
        return False


def incremecol(asheet, rows=1000, secs=0):
    arange = 'A1:B%s' % rows
    cells = asheet.range(arange)
    for cell in cells:
        col = cell.col
        row = cell.row
        #print("row: %s, col: %s" %(row, col))
        if col == 2:
            if cell.value:
                cell.value = int(cell.value) + 1
            else:
                cell.value = val
        val = cell.value
    try:
        updatecells(credentials, cells, asheet)
        return True
    except:
        return False


def chunkulate(asheet, rows=1000, stepby=50, secs=0):
    # Batch-copy incremented integer to next column. Use gs.copycolumn() to set up.
    chunks = [(x+1, x+stepby) for x in list(range(rows)) if x%stepby == 0]
    for chunkdex, atuple in enumerate(chunks):
        arange = 'B%s:C%s' % atuple
        cells = asheet.range(arange)
        for cell in cells:
            col = cell.col
            row = cell.row
            if col == 3:
                if cell.value:
                    cell.value = ''
                else:
                    cell.value = chunkdex+1
            val = cell.value
        yield True
        try:
            updatecells(credentials, cells, asheet)
        except Exception as e:
            import traceback
            print(traceback.print_exc())
            return False


def shoveit(asheet, rows=1000):
    import shelve
    arange = 'A1:D%s' % rows
    cells = asheet.range(arange)
    localdb = shelve.open(asheet.title)
    rowmem = []
    for cell in cells:
        col = cell.col
        row = cell.row
        if col == 4:
            cell.value = "stuffed"
            localdb[str(row)] = rowmem
            rowmem = []
        else:
            rowmem.append(cell.value)
    localdb.close()
    try:
        updatecells(credentials, cells, asheet)
        return True
    except:
        return False


def yankit(asheet, rows=1000):
    import shelve
    arange = 'D1:D%s' % rows
    cells = asheet.range(arange)
    localdb = shelve.open(asheet.title)
    for cell in cells:
        row = cell.row
        cell.value = localdb[str(row)]
    localdb.close()
    try:
        updatecells(credentials, cells, asheet)
        return True
    except:
        return False


scopes = ["https://www.googleapis.com/auth/analytics.readonly", 
          "https://www.googleapis.com/auth/webmasters.readonly", 
          "https://spreadsheets.google.com/feeds/",
          "https://www.googleapis.com/auth/gmail.modify"]


path = os.path.dirname(os.path.realpath('__file__'))
filename = '%s/oauth.dat' % path
flow = OAuth2WebServerFlow(client_id, client_secret, scopes, 
                           redirect_uri='urn:ietf:wg:oauth:2.0:oob',
                           response_type='code',
                           approval_prompt='force',
                           access_type='offline')
authorize_url = flow.step1_get_authorize_url()
storage = file.Storage(filename)
credentials = storage.get()
argparser = argparse.ArgumentParser(add_help=False)
parents = [argparser]
parent_parsers = [tools.argparser]
parent_parsers.extend(parents)
parser = argparse.ArgumentParser(
    description="__doc__",
    formatter_class=argparse.RawDescriptionHelpFormatter,
    parents=parent_parsers)
flags = parser.parse_args(['--noauth_local_webserver'])
try:
    http = credentials.authorize(http = httplib2.Http())
except:
    pass
if credentials is None or credentials.invalid:
    credentials = tools.run_flow(flow, storage, flags)
else:
    credentials.refresh(http)
with open(filename) as json_file:
    jdata = json.load(json_file)
token = jdata['access_token']
creds = MyCreds(access_token=token)
connect = gspread.authorize(creds)

# Force IPython Notebook to not buffer output
sys.stdout = Unbuffered(sys.stdout)

