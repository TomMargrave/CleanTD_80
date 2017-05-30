# CleanTD_80.vbs
VBScript CleanTD_80.vbs is used to assist in clearing out HPE ALM temporary directory.  The folder is located under %temp$\TD_80.  This folder holds all of the test and data locally on client machine to allow the client machine to run or perform operations. Currently there is no method clean up old or un-used files from the TD_80 Folder.

## Purpose

Currently there is no method clean up old or un-used files from the TD_80 Folder.  This script is an example of how to clean the directory and can be used with Windows scheduler.

#### Issues addressed

1. Clean un-used files older than x number of days (currently 14 days).
2. If this is scheduled it could cause issue with automated run.  The script detects if UFT is running and waits for it to complete.
3. ALM tries to connect to client during cleaning. This sets the registry so that machine can not be accessed during cleaning.

## Overview of Code steps
Runing of this code can be done by clicking on the VBS file or scheduling with Windows task scheduler.

    Note: If the company's firewall policy does not allow downloading of VBS files,
    copy the  text and place it into Notepad then save the file with .vbs extension.



### Usage:

usage:   CleanTD_80.vbs


### Thank you
Thank you to Michael Deveaux for testing and reviewing code.
