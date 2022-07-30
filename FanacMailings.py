from __future__ import annotations

from dataclasses import dataclass
import csv
import os
import re
import datetime

from Settings import Settings
from HelpersPackage import FindAndReplaceBracketedText, ParseFirstStringBracketedText, SortMessyNumber, NormalizePersonsName, Int0, DateMonthYear, FormatLink2
from Log import LogError, LogDisplayErrorsIfAny, LogOpen


# Function to find the index of a column header
def ColIndex(headers: [str], header: str) -> int:
    if header not in headers:
        return -1
    return headers.index(header)

def main():
    LogOpen("log.txt", "log-ERRORS.txt")
    if not Settings().Load("FanacMailings settings.txt", MustExist=True, SuppressMessageBox=True):
        LogError("Could not find settings file 'FanacMailings settings.txt'")
        return

    # ---------------
    # Get the list of known apas
    # Mailings is a dictionary indexed by the apa name.
    #   The value is a dictionary indexed by the mailing number as a string
    #       The valur of *that* is a MailingDev
    mailingsInfoTable: dict[str, dict[str, MailingDev]]={}
    knownApas=Settings().Get("Known APAs")
    if len(knownApas) == 0:
        LogError("Settings file 'FanacMailings settings.txt' does not contain a value for 'Known APAs' (the list of APAs we care about here)")
        return
    knownApas=[x.replace('"', '').strip() for x in knownApas.split(",")]

    # ---------------
    # for each known apa, read Joe's APA mailings data if it exists
    for apaName in knownApas:
        csvname=apaName+".csv"
        # Skip missing csv files
        if not os.path.exists(apaName+".csv"):
            continue
        # Read the csv file
        try:
            with open(csvname, 'r') as csvfile:
                # Read it into a list of lists
                filereader=csv.reader(csvfile, delimiter=',', quotechar='"')
                mailingsdata=[x for x in filereader]
        except FileNotFoundError:
            LogError(f"Could not open CVS file {csvfile}")
            return
        # Separate out the header row
        mailingsheaders=mailingsdata[0]
        mailingsdata=mailingsdata[1:]

        issueCol=ColIndex(mailingsheaders, "Issue")
        if issueCol == -1:
            LogError(f"{csvname} does not contain an 'Issue' column")
            return
        monthCol=ColIndex(mailingsheaders, "Month")
        if monthCol == -1:
            LogError(f"{csvname} does not contain an 'Month' column")
            return
        yearCol=ColIndex(mailingsheaders, "Year")
        if yearCol == -1:
            LogError(f"{csvname} does not contain an 'Year' column")
            return
        editorCol=ColIndex(mailingsheaders, "Editor")
        if editorCol == -1:
            LogError(f"{csvname} does not contain an 'Editor' column")
            return

        # Create a dictionary of mailings for this APA.  Indexed by the mailing number as a string.
        @dataclass
        class MailingDev:
            Number: str=""
            Year: str=""
            Month: str=""
            Editor: str="Editor?"

        mailingsInfoTable[apaName]={}
        for md in mailingsdata:
            mailingsInfoTable[apaName][md[1]]=MailingDev(Number=md[issueCol], Year=md[yearCol], Month=md[monthCol], Editor=md[editorCol])

    # ---------------
    # Get the CSV source file (generated by FanacAnalyzer) out of settings and open it
    sourceCVSfile=Settings().Get("CSVSource")
    if len(sourceCVSfile) == 0:
        LogError("Settings file 'FanacMailings settings.txt' does not contain a value for CSVSource (the file generated by FanacAnalyzer)")
        return

    try:
        with open(sourceCVSfile, 'r') as csvfile:
            filereader=csv.reader(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)
            mailingsdata=[x for x in filereader]
    except FileNotFoundError:
        LogError(f"Could not open CVS file {sourceCVSfile}")
        return

    mailingsheaders=mailingsdata[0]
    mailingscol=ColIndex(mailingsheaders, "Mailings")
    if mailingscol == -1:
        LogError(f"Could not find a mailings column in {mailingsheaders}")
        return

    # ---------------------------
    # Turn the data from FanacAnalyzer into a dictionary of the form dict(apa, dict(mailing, data))
    apas: dict[str, dict[str, [str]]]={}
    for row in mailingsdata:
        # The mailings column is of the form   ['FAPA 20 & VAPA 23']
        mailings=row[mailingscol]
        mailings=mailings.removeprefix("['").removesuffix("']")
        mailings=[x.strip() for x in mailings.split("&")]
        for mailing in mailings:
            for apa in knownApas:
                m=re.match(f"{apa}\s(.*)$", mailing)
                if m is not None:
                    if apa not in apas.keys():
                        apas[apa]={}
                    mailingnumber=m.groups()[0]
                    if mailingnumber not in apas[apa].keys():
                        apas[apa][mailingnumber]=[]
                    apas[apa][m.groups()[0]].append(row)

    # ------------------
    # We've slurped in all the data.  Now create the index files for each issue
    # We will create a file in the ReportsDir for each APA, and put the individual issue index pages there
    reportsdir=Settings().Get("ReportsDir")
    if len(reportsdir) == 0:
        LogError("Settings file 'FanacMailings settings.txt' does not contain a value for ReportsDir (the directory to be used for reports)")
        return
    if not os.path.exists(reportsdir):
        os.mkdir(reportsdir)

    # Read the mailing template file
    templateFilename=Settings().Get("Template-Mailing")
    if len(templateFilename) == 0:
        LogError("Settings file 'FanacMailings settings.txt' does not contain a value for Template-Mailing (the template for an individual mailing page)")
        return
    try:
        with open(templateFilename, "r") as file:
            templateMailing="".join(file.readlines())
    except FileNotFoundError:
        LogError(f"Could not open the mailing template file, '{templateFilename}'")
        return

    # Walk through the info from FanacAnalyzer.
    # For each APA that we found there:
    #   Create an apa HTML page listing (and linking to) all the mailing pages
    #   Create all the individual mailing pages
    for apa in apas.keys():

        # Make sure that a directory exists for that APA's html files
        if not os.path.exists(os.path.join(reportsdir, apa)):
            os.mkdir(os.path.join(reportsdir, apa))

        mailingInfo={}
        if apa in mailingsInfoTable:
            mailingInfo=mailingsInfoTable[apa]

        # For each mailing of that APA generate a mailing page.
        # Also accumulate the info needed to produce the apa page
        listOfMailings=[]
        for mailing in apas[apa]:

            # Do a mailing page

            # First, the top matter
            # <div><fanac-top>
            # <table class=topmatter>
            # <tr><td class=topmatter>mailing</td></tr>
            # <tr><td class=topmatter>editor</td></tr>
            # <tr><td class=topmatter>date</td></tr>
            # <tr><td class=topmatter>APA Mailing</td></tr>
            # </table>
            # </fanac-top></div>
            mailingPage=templateMailing
            start, mid, end=ParseFirstStringBracketedText(mailingPage, "fanac-top")
            mid=mid.replace("mailing", f"{apa} mailing {mailing}")
            editor="editor?"
            when="when?"
            if mailing in mailingInfo:
                m=mailingInfo[mailing]
                editor=f"OE: {NormalizePersonsName(m.Editor)}"
                when=DateMonthYear(Int0(m.Month), Int0(m.Year))
            mid=mid.replace("editor", editor)
            mid=mid.replace("date", when)
            mailingPage=start+mid+end

            # Now add the necessary header info
            start, mid, end=ParseFirstStringBracketedText(mailingPage, "fanac-title")
            mid=mid.replace("name of mailing", f"{apa}-{mailing}")
            mailingPage=start+mid+end

            start, mid, end=ParseFirstStringBracketedText(mailingPage, "head")
            mid=mid.replace("mailing content", f"{mailing}, {editor}, {when}, {apa}-mailing")
            mailingPage=start+mid+end

            # Now the bottom matter (the list of fanzines)
            newtable="<tr>\n"
            # Generate the header row, selecting only those headers which are in this dict:
            colSelectionAndOrder=["IssueName", "Editor", "PageCount"]   # The columns to be displayed in order
            colNaming=["Contribution", "Editor", "Page Count"]      # The corresponding column names
            linkCol=mailingsheaders.index("IssueName")   # The column to have the link to the issue
            seriesUrlCol=mailingsheaders.index("DirURL")   # The column containing the full URL of the issue's directory'
            issueUrlCol=mailingsheaders.index("PageName")   # The column containing the issue's URL (just the issue, not the issue's directory)
            colsSelected=[]     # Retain the indexes of the selected headers to generate the table rows
            for col in range(len(colSelectionAndOrder)):
                if colSelectionAndOrder[col] not in mailingsheaders:
                    LogError(f"Can't find a column named '{colSelectionAndOrder[col]}' the headers list: [{mailingsheaders}]")
                    return
                newtable+=f"<th>{colNaming[col]}</th>\n"
                colsSelected.append(mailingsheaders.index(colSelectionAndOrder[col]))
            newtable+="</tr>\n"

            # Now generate the data rows in the mailings table
            for row in apas[apa][mailing]:
                newtable+="<tr>\n"
                for col in colsSelected:
                    if col == linkCol:
                        fullUrl=row[seriesUrlCol]+"/"+row[issueUrlCol]
                        newtable+=f"<th><a href={fullUrl}>{row[col]}</a></th>\n"
                    else:
                        newtable+=f"<th>{row[col]}</th>\n"
                newtable+="</tr>\n"
            newtable=newtable.replace("\\", "/")

            # Insert the new issues table into the template
            mailingPage, success=FindAndReplaceBracketedText(mailingPage, "fanac-rows", newtable)
            if not success:
                LogError("Could not add issues table to mailing page at 'fanac-rows'")
                return

            # Add the updated date/time
            mailingPage, success=FindAndReplaceBracketedText(mailingPage, "fanac-updated", f"Updated {datetime.datetime.now().strftime('%m/%d/%Y, %H:%M:%S')}")

            # Write the mailing file
            with open(os.path.join(reportsdir, apa, mailing)+".html", "w") as file:
                mailingPage=mailingPage.split("/n")
                file.writelines(mailingPage)

            # Also append to the accumulator for the apa page
            listOfMailings.append((mailing, None))  # To be expanded

        #------------------------------
        # Now that the mailing pages are all done, do an apa page
        # Read the apa template file
        templateFilename=Settings().Get("Template-APA")
        if len(templateFilename) == 0:
            LogError("Settings file 'FanacMailings settings.txt' does not contain a value for Template-APA (the template for an APA page)")
            return
        try:
            with open(templateFilename, "r") as file:
                templateApa="".join(file.readlines())
        except FileNotFoundError:
            LogError(f"Could not open the APA template file, '{templateFilename}'")
            return
        loc=templateApa.find("</fanac-rows>")
        if loc < 0:
            LogError(f"The APA template '{templateFilename}' is missing the '</fanac-rows>' indicator.")
            return
        templateApaFront=templateApa[:loc]
        templateApaRear=templateApa[loc+len("</fanac-rows>"):]

        # Now sort the accumulation of mailings ito numerical order and create the apa page
        listOfMailings=sorted(listOfMailings, key=lambda x: SortMessyNumber(x[0]))
        for mailingTuple in listOfMailings:
            mailing=mailingTuple[0]
            editor="editor?"
            when="when?"
            if mailing in mailingInfo:
                m=mailingInfo[mailing]
                editor=NormalizePersonsName(m.Editor)
                when=DateMonthYear(Int0(m.Month), Int0(m.Year))

            templateApaFront+=f"\n<tr><td><a href={mailing}.html>{mailing}</a></td><td>{when}</td><td>{editor}</td></tr>"

        # Add the updated date/time
        templateApaRear, success=FindAndReplaceBracketedText(templateApaRear, "fanac-updated", f"Updated {datetime.datetime.now().strftime('%m/%d/%Y, %H:%M:%S')}>")

        # Write out the APA list of all mailings
        with open(os.path.join(reportsdir, apa, "index.html"), "w") as file:
            file.writelines(templateApaFront+templateApaRear)


# Run main()
if __name__ == "__main__":
    main()
    LogDisplayErrorsIfAny()


