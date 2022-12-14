from __future__ import annotations

from dataclasses import dataclass
import csv
import os
import re
import datetime

from FanzineIssueSpecPackage import FanzineDate

from Settings import Settings
from HelpersPackage import FindAndReplaceBracketedText, ParseFirstStringBracketedText, SortMessyNumber, SortTitle, NormalizePersonsName, Int0, DateMonthYear
from HelpersPackage import FindIndexOfStringInList
from Log import LogError, Log, LogDisplayErrorsIfAny, LogOpen


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
            LogError(f"Could not open CSV file {csvname}")
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
        class MailingDev:

            def __init__(self, Number: str="", Year: str="", Month: str="", Editor: str=""):
                self.Number: str=Number
                self.Editor: str=Editor

                fd=FanzineDate()
                if Month != "":
                    fd.Month=Month
                if Year != "":
                    fd.Year=Year
                self.Date: FanzineDate=fd

            @property
            def Year(self) -> int:
                return self.Date.Year
            @Year.setter
            def Year(self, y: int) -> None:
                self.Date.Year=y

            @property
            def Month(self) -> int:
                return self.Date.Month
            @Month.setter
            def Month(self, m: int) -> None:
                self.Date.Month=m

        mailingsInfoTable[apaName]={}
        for md in mailingsdata:
            mailingsInfoTable[apaName][md[1]]=MailingDev(Number=md[issueCol], Year=md[yearCol], Month=md[monthCol], Editor=md[editorCol])


    # ---------------
    # Get the location of the CSV source file (generated by FanacAnalyzer) out of settings and open it
    sourceCSVfile=Settings().Get("CSVSource")
    if len(sourceCSVfile) == 0:
        LogError("Settings file 'FanacMailings settings.txt' does not contain a value for CSVSource (the file generated by FanacAnalyzer)")
        return

    try:
        with open(sourceCSVfile, 'r') as csvfile:
            filereader=csv.reader(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)
            mailingsdata=[x for x in filereader]
    except FileNotFoundError:
        LogError(f"Could not open CSV file {sourceCSVfile}")
        return

    if len(mailingsdata) < 100:
        LogError(f"There are {len(mailingsdata)} items in {sourceCSVfile} -- there should be hundreds")
        return

    mailingsheaders=mailingsdata[0]
    mailingscol=ColIndex(mailingsheaders, "Mailings")
    if mailingscol == -1:
        LogError(f"Could not find a 'Mailings' column in {mailingsheaders}")
        LogError(f"The column headers found were {mailingsheaders}")
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
        LogError("Settings file 'FanacMailings settings.txt' does not contain a value for Template-Mailing (the name of the template file for an individual mailing page)")
        return
    try:
        with open(templateFilename, "r") as file:
            templateMailing="".join(file.readlines())
    except FileNotFoundError:
        LogError(f"Could not open the mailing template file: '{templateFilename}'")
        return

    # All the pages we generate here need the same kinds of information to be added:
    #   Page title
    #   Page metadata
    #   Updated timestamp
    def AddBoilerplate(page: str, title: str, metadata: str) -> str:
        start, mid, end=ParseFirstStringBracketedText(page, "fanac-title")
        mid=mid.replace("title of page", title)
        page=start+mid+end

        start, mid, end=ParseFirstStringBracketedText(page, "head")
        mid=mid.replace("mailing content", metadata)
        page=start+mid+end

        # Add the updated date/time
        page, success=FindAndReplaceBracketedText(page, "fanac-updated", f"Updated {datetime.datetime.now().strftime('%m/%d/%Y, %H:%M:%S')}")
        return page

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
            # </table>
            # </fanac-top></div>
            mailingPage=templateMailing
            start, mid, end=ParseFirstStringBracketedText(mailingPage, "fanac-top")
            editor=""   #"editor?"
            when="" #"when?"
            if mailing in mailingInfo:
                m=mailingInfo[mailing]
                editor=f"OE: {NormalizePersonsName(m.Editor)}"
                when=m.Date.FormatDate("%B %Y")
                number=m.Number
            else:
                number=mailing
            mid=mid.replace("editor", editor)
            mid=mid.replace("date", when)
            mid=mid.replace("mailing", f"{apa} Mailing #{number}")
            mailingPage=start+mid+end

            mailingPage=AddBoilerplate(mailingPage, f"{apa}-{mailing}", f"{mailing}, {editor}, {when}, {apa}-mailing")

            # Now the bottom matter (the list of fanzines)
            newtable="<tr>\n"
            # Generate the header row, selecting only those headers which are in this dict:
            colSelectionAndOrder=["IssueName", "Editor", "PageCount"]   # The columns to be displayed in order
            colNaming=["Contribution", "Editor", "Pages"]      # The corresponding column names
            linkCol=mailingsheaders.index("IssueName")   # The column to have the link to the issue
            seriesUrlCol=mailingsheaders.index("DirURL")   # The column containing the full URL of the issue's directory'
            issueUrlCol=mailingsheaders.index("PageName")   # The column containing the issue's URL (just the issue, not the issue's directory)
            pagesCol=mailingsheaders.index(("PageCount"))
            colsSelected=[]     # Retain the indexes of the selected headers to generate the table rows
            for col in range(len(colSelectionAndOrder)):
                colindex=FindIndexOfStringInList(mailingsheaders, colSelectionAndOrder[col])
                if colindex is None:
                    LogError(f"Can't find a column named '{colSelectionAndOrder[col]}' in the headers list: [{mailingsheaders}]")
                    return
                newtable+=f"<th>{colNaming[col]}</th>\n"
                colsSelected.append(colindex)
            newtable+="</tr>\n"

            # Sort the contributions into order by fanzine name
            apas[apa][mailing]=sorted(apas[apa][mailing], key=lambda x: SortTitle(x[0]))

            # Now generate the data rows in the mailings table
            for row in apas[apa][mailing]:
                newtable+="<tr>\n"
                for col in colsSelected:
                    if col == linkCol:
                        fullUrl=row[seriesUrlCol]+"/"+row[issueUrlCol]
                        newtable+=f"<td><a href={fullUrl}>{row[col]}</a></td>\n"
                    elif col == pagesCol:
                        newtable+=f"<td style='text-align: right'>{row[col]}&nbsp;&nbsp;</td>"
                    else:
                        newtable+=f"<td>{row[col]}</td>\n"
                newtable+="</tr>\n"
            newtable=newtable.replace("\\", "/")

            # Insert the new issues table into the template
            mailingPage, success=FindAndReplaceBracketedText(mailingPage, "fanac-rows", newtable)
            if not success:
                LogError("Could not add issues table to mailing page at 'fanac-rows'")
                return

            # Insert the label for the button taking you up one level to all mailings for this APA
            mailingPage, success=FindAndReplaceBracketedText(mailingPage, "fanac-AllMailings", f"All {apa} mailings")
            if not success:
                LogError("Could not change button text on mailing page at 'fanac-AllMailings'")
                return

            # Modify the Mailto: so that the page name appears as the subject
            mailingPage, success=FindAndReplaceBracketedText(mailingPage, "fanac-ThisPageName", f"{apa}:{mailing}")
            if not success:
                LogError("Could not change mailto Subject on mailing page at 'fanac-ThisPageName'")
                return

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

        # Add the APA's name at the top
        start, mid, end=ParseFirstStringBracketedText(templateApa, "fanac-top")
        mid=mid.replace("apa-name", apa)
        templateApa=start+mid+end

        # Add random descriptive information if a file <apa>-bumpf.txt exists.  (E.g., SAPS-bumpf.txt)
        fname=apa+"-bumpf.txt"
        if os.path.exists(fname):
            with open(fname, "r") as file:
                bumpf=file.read()
            if len(bumpf) > 0:
                start, mid, end=ParseFirstStringBracketedText(templateApa, "fanac-bumpf")
                if len(end) > 0:
                    mid=bumpf+"<p>"
                    templateApa=start+mid+end
            Log(f"Bumpf added to {apa} page")
        else:
            Log(f" No {fname} file found, so no bumpf added to {apa} page.")

        templateApa=AddBoilerplate(templateApa, f"{apa} Mailings", f"{apa} mailings")

        loc=templateApa.find("</fanac-rows>")
        if loc < 0:
            LogError(f"The APA template '{templateFilename}' is missing the '</fanac-rows>' indicator.")
            return
        templateApaFront=templateApa[:loc]
        templateApaRear=templateApa[loc+len("</fanac-rows>"):]


        # Now sort the accumulation of mailings into numerical order and create the apa page
        listOfMailings=sorted(listOfMailings, key=lambda x: SortMessyNumber(x[0]))
        for mailingTuple in listOfMailings:
            mailing=mailingTuple[0]
            editor=""   #"editor?"
            when="" #"when?"
            if mailing in mailingInfo:
                m=mailingInfo[mailing]
                editor=NormalizePersonsName(m.Editor)
                when=DateMonthYear(Int0(m.Month), Int0(m.Year))

            templateApaFront+=f"\n<tr><td><a href={mailing}.html>{mailing}</a></td><td>{when}</td><td>{editor}</td><td style='text-align: right'>{len(apas[apa][mailing])}&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>"

        templateApa=templateApaFront+templateApaRear

        # Add counts of mailings and contributions to bottom
        start, mid, end=ParseFirstStringBracketedText(templateApa, "fanac-totals")
        numConts=0
        for mailingTuple in listOfMailings:
            numConts+=len(apas[apa][mailingTuple[0]])

        mid=f"{len(listOfMailings)} mailings containing {numConts} individual contributions"
        templateApa=start+mid+end

        # Add the updated date/time
        templateApa, success=FindAndReplaceBracketedText(templateApa, "fanac-updated", f"Updated {datetime.datetime.now().strftime('%m/%d/%Y, %H:%M:%S')}>")

        # Make the mailto correctly list the apa in the subject line
        templateApa, success=FindAndReplaceBracketedText(templateApa, "fanac-APAPageMailto", f"Issue related to APA {apa}")
        if not success:
            LogError(f"The APA template '{templateFilename}' is missing the '</fanac-APAPageMailto>' indicator.")
            return

        # Write out the APA list of all mailings
        with open(os.path.join(reportsdir, apa, "index.html"), "w") as file:
            file.writelines(templateApa)

    #--------------------------------------------
    # Generate the All Apas root page
    templateFilename=Settings().Get("Template-allAPAs")
    if len(templateFilename) == 0:
        LogError("Settings file 'FanacMailings settings.txt' does not contain a value for Template-allAPAs (the template for the page listing all APAs)")
        return
    try:
        with open(templateFilename, "r") as file:
            templateAllApas="".join(file.readlines())
    except FileNotFoundError:
        LogError(f"Could not open the all APAs template file, '{templateFilename}'")
        return

    templateAllApas=AddBoilerplate(templateAllApas, f"Mailings for All APAs", f"Mailings for All APAs")

    listText="&nbsp;<ul>"
    for apa in apas.keys():
        listText+=f"<li><a href='{apa}/index.html'>{apa}</a></li>\n"
    listText+="</ul>\n"
    templateAllApas, success=FindAndReplaceBracketedText(templateAllApas, "fanac-list", listText)

    with open(os.path.join(reportsdir, "index.html"), "w") as file:
        file.writelines(templateAllApas)

# Run main()
if __name__ == "__main__":
    main()
    LogDisplayErrorsIfAny()


