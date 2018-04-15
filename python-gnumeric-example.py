#!/usr/bin/python
# A practical example of using Gnumeric's introspection interface
# to python.
#
# This program will create a bitcoin price tracking spreadsheet
# if it does not already exist, format it, and then access 
# Coinbase's json bitcoin price api and append the latest
# USD exchange rate to the sheet, save it and quit.
#
# It demonstrates file I/O, and many of the Gnumeric interface
# functions to read and write date in cells and apply formats.
#
# Examples of most gnumeric operations can be found at
# https://git.gnome.org/browse/gnumeric/tree/test/t3001-introspection-simple.py
# 
# Requirements:
# gnumeric 1.12.40 or better, compiled with --enable-introspection
# supporting libraries (goffice) also compiled with --enable-introspection
#
# example by Dean McCarron (dm-gnome mercuryresearch.com)

import os
import gi
gi.require_version('Gnm', '1.12')
gi.require_version('GOffice', '0.10') 
from gi.repository import Gnm
from gi.repository import GOffice

global cc
global ioc

# Setup and initialization
def gnm_init():
    global cc
    global ioc

    # Initialize Gnm itself
    Gnm.init()

    # create a stderr reporting context
    cc = Gnm.CmdContextStderr.new()

    # Load plugins
    Gnm.plugins_init(cc)

    # A context for io operations
    ioc = GOffice.IOContext.new(cc)

    return()

# Call with a filename, returns a WorkbookView object.
def wb_open(filename):
    global ioc
    
    uri = GOffice.filename_to_uri(filename)

    # Check to make sure file exists first, if it doesn't create new spreadsheet and set uri
    if os.path.isfile(filename):
        # Open file and get wbv
        wbv = Gnm.WorkbookView.new_from_uri(uri, None, ioc, None)
        wb = wbv.props.workbook
        wbv = None #Free wbv object
    else:
        wb = Gnm.Workbook.new_with_sheets(1)
        wb.props.uri = uri
    
    return(wb)

# Will save to existing filename if called with just WorkbookView object, otherwise will save to filename.
def wb_save(wb, filename=None):
    global cc

    if filename is None:
        uri = wb.props.uri
    else:
        uri = GOffice.filename_to_uri(filename)

    fs = GOffice.FileSaver.for_file_name(uri)
    wbv=Gnm.WorkbookView.new(wb)
    if not wbv.save_as(fs, uri, cc):
        raise IOError("Failed to save workbook")
    wbv = None #Free wbv object

    return()

def main():

    gnm_init()
    
    # Open (or create if it doesn't exist) our price tracking workbook.
    wb = wb_open("btcprices.gnumeric")
    
    # Select the first sheet in the workbook, check the name, and update name.
    sheet = wb.sheet_by_index(0)
    if sheet.props.name != "Bitcoin Prices":
        sheet.props.name = "Bitcoin Prices"

    # Check to see if there is a title row by checking if there's any value set at B2.
    # If none is set, populate the title row and style it. This is effectively a 
    # run-once operation.

    if sheet.cell_get_value(1,1) is None:
        # Labels
        sheet.cell_set_text(1,1,"Timestamp (UTC)")
        sheet.cell_set_text(2,1,"BTC Price (USD)")

        # Bold
        st = Gnm.Style.new()
        st.set_font_bold(1)
        r = Gnm.Range()
        # col row col row range, so 1:1 (B2)
        r.init(1,1,2,1)
        sheet.apply_style(r,st)

        # Adjust Column Width
        # TBD - no blessed direct sizing interface, using autofit
        #r = Gnm.Range()
        #r.init(0,0,4,4)
        #Gnm.colrow_autofit_col(sheet, r)

    # Find first unused row in sheet to set up the addition of data
    for row in range(2, sheet.props.rows):
        if sheet.cell_get_value(1,row) is None:
            break

    # At this point, row points to the next open row

    # We will use the json interface to coinbase at https://api.coindesk.com/v1/bpi/currentprice.json
    # to extract the updated time and the current USD value, and then populate the row with that data.
    # Note that the user-agent needs to be set to a sane value, as urllib's default is banned due to abuse.
    # (per coinbase's request in their json feed, please don't abuse the interface with excessive requests!)
    #

    import urllib2, json

    url   = "https://api.coindesk.com/v1/bpi/currentprice.json"
    agent = "Mozilla/5.0 (Windows; U; Windows NT 5.1; it; rv:1.8.1.11) Gecko/20071127 Firefox/2.0.0.11"

    opener = urllib2.build_opener()
    opener.addheaders = [('User-Agent', agent)]
    response = opener.open(url)
    data = json.loads(response.read())

    # Update the row with the time the price was updated
    # we're getting the time in ISO 8601 format, which gnumeric won't directly
    # interpret, so we're going to convert it to python time, and the re-write it
    # in yyyy-mm-dd hh:mm:ss format that gnumeric is fine with.

    import dateutil.parser
    import time 

    tdate = dateutil.parser.parse( data['time']['updated'] )

    # Note we use set_text and not set_value(Gnm.Value.new_string()) because we
    # want gnumeric to interpret our input for us into a time value. If we just
    # wanted the raw text, we'd use set value and new_string.

    sheet.cell_set_text(1,row, tdate.strftime("%Y-%m-%d %H:%M:%S"))

    # ... then we update with the current USD exchange rate.
    # We get a little fancy using "locale" to deal with comma separateors in the 
    # ascii to floating point conversion using atof. (We could just set text in gnumeric as well, but
    # then we wouldn't be showing off setting floating point values ;-)

    import locale
    locale.setlocale(locale.LC_NUMERIC, '')
    sheet.cell_set_value(2,row, Gnm.Value.new_float( locale.atof(data['bpi']['USD']['rate'])))


    # Let's format the price to just dollars with a thousands separator
    st = Gnm.Style.new()
    st.set_format_text("$0,000")
    r = Gnm.Range()
    r.init(2,row,2,row)
    sheet.apply_style(r,st)


    # Whew! All done. Save it out and quit.

    wb_save(wb)

    # Cleanup
    wb = None
    ioc = None

    quit()

# This allows this file to be included elsewhere
if __name__ == "__main__":
    # execute only if run as a script
    main()

