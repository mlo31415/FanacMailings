from __future__ import annotations

import csv
import os
import re

from Settings import Settings
from HelpersPackage import FindAndReplaceBracketedText
from Log import Log, LogError, LogDisplayErrorsIfAny

def main():
    if not Settings().Load("FanacMailings settings.txt"):
        LogError("Could not find settings file 'FanacMailings settings.txt'")
        return

    # Get the CSV sourced file out of settings and open it
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

    # Function to find the index of a column header
    def ColIndex(headers: [str], header: str) -> int:
        if header not in headers:
            return -1
        return headers.index(header)

    mailingsheaders=mailingsdata[0]
    mailingscol=ColIndex(mailingsheaders, "Mailings")
    if mailingscol == -1:
        LogError(f"Could not find a mailings column in {mailingsheaders}")
        return
    knownApas=Settings().Get("Known APAs")
    knownApas=[x.replace('"', '').strip() for x in knownApas.split(",")]

    apas: dict[str, dict[str, [str]]]={}
    for row in mailingsdata:
        # The mailings column is of the form
                        # ['FAPA 20 & VAPA 23']
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
                    i=0


    # We've slurped in all the data.  Now create the index files for each issue
    # We will create a file in the ReportsDir for each APA, and put the individual issue index pages there
    reportsdir=Settings().Get("ReportsDir")
    if not os.path.exists(reportsdir):
        os.mkdir(reportsdir)

    # Read the mailing and template files
    templateFilename=Settings().Get("Template-Mailing")
    if len(templateFilename) == 0:
        LogError("Settings file 'FanacMailings settings.txt' does not contain a value for Template-Mailing (the template for an individual mailing page)")
        return
    try:
        with open(templateFilename, "r") as file:
            templateMailing="".join(file.readlines())
    except FileNotFoundError:
        LogError(f"Could not open '{templateFilename}'")
        return

    templateFilename=Settings().Get("Template-APA")
    if len(templateFilename) == 0:
        LogError("Settings file 'FanacMailings settings.txt' does not contain a value for Template-APA (the template for an APA page)")
        return
    try:
        with open(templateFilename, "r") as file:
            templateApa="".join(file.readlines())
    except FileNotFoundError:
        LogError(f"Could not open '{templateFilename}'")
        return

    # For each APA
    for apa in apas.keys():
        # Make sure that a directory exists for that APA
        if not os.path.exists(os.path.join(reportsdir, apa)):
            os.mkdir(os.path.join(reportsdir, apa))
        # For each mailing of that APA
        for mailing in apas[apa]:
            newtable="<tr>\n"
            for header in mailingsheaders:
                newtable+=f"<th>{header}</th>\n"
            newtable+="</tr>\n"
            for row in apas[apa][mailing]:
                newtable+="<tr>\n"
                for cell in row:
                    newtable+=f"<th>{cell}</th>\n"
                newtable+="</tr>\n"
            newtable=newtable.replace("\\", "/")
            issueindex=templateMailing    # Make a copy of the template
            issueindex, success=FindAndReplaceBracketedText(issueindex, "fanac-rows", newtable)
            if success:
                    with open (os.path.join(reportsdir, apa, mailing)+".html", "w") as file:
                        issueindex=issueindex.split("/n")
                        file.writelines(issueindex)
            i=0

    i=0


# Run main()
if __name__ == "__main__":
    main()
    LogDisplayErrorsIfAny()