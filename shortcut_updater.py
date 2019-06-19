#! python3

# Manar Naboulsi - 26 March, 2018
#
# Document Updates Checker - Checks to see whether SDS, UAeCRFs, eCRFs, DMPs, or CSPs 
#                            have been updated.  If so, a shortcut is created and an
#                            email is sent to programming team.                             

# Arguments:
#   - location = filepath of shortcuts


import logging
import os
import re
import smtplib
import win32com.client
from datetime import datetime


# lists studies to overlook
oldstudies = ['SM04755-PSO-01', 'SM04690-OA-01', 'SM04690-OA-02', 'SM04690-OA-04', 
              'SM04755-ONC-01', 'SM04646-IPF-01', 'SM04646-IPF-02', 'SM04646-IPF-03']

# doc-specific info 
doctype = ['DMP', 'eCRF CG', 'SDS', 'UAeCRF']
doctype1 = ['DMP', 'eCRF', 'SDS', 'UAeCRF']
doctype_med = dict(DMP='DMP', eCRF='eCRF CG', SDS='SDS', UAeCRF='UAeCRF')
doctype_long = dict(DMP='-data-management-plan-', eCRF='-ecrf-completion-guidelines-', 
                    SDS='-study-design-specs-', UAeCRF='-unique-annotated-ecrfs-')
doctype_ext = dict(DMP='\.pdf$', eCRF='\.pdf$', SDS='\.xlsx$', UAeCRF='\.pdf$')
doctype_ext1 = dict(DMP='PDF', eCRF='PDF', SDS='XLSX', UAeCRF='PDF')


def shortcut_updater(location):
    # initializes logger
    logging.basicConfig(level=logging.DEBUG, filename='H:\\logs\\doc_updates_checker_log.log')
    logging.info('-----------------------------------------------------------------------------')
    logging.info('Starting at ' + str(datetime.now()))
    
    # initializes email message with legend
    message = '  ~ :  Updated\n** :  Action Needed\n'
   
    # looks through shortcut folder
    for study in os.listdir(location):
        studypath = os.path.join(location, study)
        
        doclink = dict.fromkeys(doctype_long)  # creates dict with doctypes as keys to track whether shortcut is found
        # gets study folders from location and overlooks old studies
        if study not in oldstudies and os.path.isdir(studypath) is True:
            message += '\n' + study
            
            # gets shortcuts from folders
            for link in os.listdir(studypath):
                scpath_cur = os.path.join(studypath, link)

                # looks through shortcut files
                if scpath_cur.endswith('.lnk') is True:
                    
                    # gets real path from shortcut
                    shell = win32com.client.Dispatch('WScript.Shell')
                    sc_cur = shell.CreateShortcut(scpath_cur)
                    realpath = sc_cur.Targetpath
        
                    for i in range(len(doctype)):
                        if doctype[i] in link:
                            doclink[doctype1[i]] = True
                            
                            # checks that shortcut is not corrupted
                            try:
                                open(realpath)
                            except:
                                message += '\n  ** ' + doctype[i] + ': ' + os.path.basename(realpath) + ' cannot be opened.' \
                                           + '\n             Please ensure shortcut does not point to a folder and is not corrupted.' \
                                           + '\n             If the shortcut points to a file and is not corrupted, please delete the original shortcut and' \
                                           + '\n             create a new one, as it is pointing to a document that has been renamed.'
                                continue
                            
                            # regex
                            pattern = "".join([
                                    '^(sm\d{5}-[a-z]+?-\d{2})',   # study number       group 1: study number
                                    doctype_long[doctype1[i]],    # document
                                    'v(\d+)-(\d+)',               # version number     group 2-3: version number
                                    '-(\d{4})([a-z]{3})(\d{2})',  # version date       group 4-6: yyyymmmdd
                                    doctype_ext[doctype1[i]]])    # ext
                            regex = re.compile(pattern, re.VERBOSE)

                            # finds version of current shortcut
                            curmatch = re.match(regex, os.path.basename(realpath).lower())
                            cur_version = 0                            

                            # if shortcut points to incorrectly named file
                            if curmatch is None:
                                message += '\n  ** ' + doctype[i] + ': No ' + doctype[i] + ' found. \n             Shortcut points to "' + os.path.basename(realpath) \
                        			       + '\n             Please ensure shortcut, naming convention, and filetype (' + doctype_ext1[doctype1[i]] + ') are correct.' \
                        			       + '\n             e.g. "ab1234-xx-01' + doctype_long[doctype1[i]] + 'v1-0-2019feb08.' + doctype_ext1[doctype1[i]].lower() + '"'\
                                           + '\n             If you rename the target document, please delete the original shortcut and create a new one.'

                            # if shortcut successfully points to file
                            else:
                                cur_version = float(curmatch.group(2) + '.' + curmatch.group(3))
                        
                                new_version = ''
                                # finds other docs in folder doc is saved in
                                for item in os.listdir(os.path.dirname(realpath)):
                                    othmatch = re.match(regex, item)
                                    if othmatch is not None:
                                        version = float(othmatch.group(2) + '.' + othmatch.group(3))
                                        if version > cur_version:
                                            new_version = item

                                # if there is an updated doc, shortcut is replaced
                                if new_version != '':
                                    newmatch = re.match(regex, new_version)
                                    # deletes old shortcut
                                    os.remove(scpath_cur)
                                    message += '\n   ~ ' + doctype1[i] + ': Deleted shortcut to ' + os.path.basename(realpath)
                                    
                                    # creates new shortcut
                                    scpath_new = os.path.join(os.path.dirname(scpath_cur), '.' + study + '_' + doctype[i] + '_V' + newmatch.group(2) + '.' + newmatch.group(3) + '.pdf.lnk')
                                    shell = win32com.client.Dispatch("WScript.Shell")
                                    sc_new = shell.CreateShortcut(scpath_new)
                                    sc_new.Targetpath = os.path.join(os.path.dirname(realpath), new_version)
                                    sc_new.save()
                                    message += '\n             Added shortcut to ' + new_version
                                else:
                                    message += '\n       ' + doctype[i] + ': Up to date. (Version ' + str(cur_version) + ')'

            # if no shortcut is found
            for key in doclink:
                if doclink[key] is None:
                    message += '\n  ** ' + doctype_med[key] + ': No shortcut found.' \
                                + '\n             Please create a shortcut or ensure an existing shortcut is spelled correctly.' \
                                + '\n             i.e. "' + doctype_med[key] + '"'

            message += '\n'   

    # preps email
    from_addr = ''  # sender email
    to_addr = []    # receiver emails
    subject = 'Weekly Study Shortcut Summary'
    body = 'Subject: {}\n\n{}'.format(subject, message)  
    
    # sends email
    try:
        smtpObj = smtplib.SMTP()  # server
        smtpObj.sendmail(from_addr, to_addr, body)
        logging.info('Email sent successfully.')
    except Exception as e:
        pass
        logging.info(e)
        logging.info('Error: Unable to send email.')
            
    # ends logger
    logging.info('Ending at ' + str(datetime.now()))    
    logging.shutdown()                                  

   



