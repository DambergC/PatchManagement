# Manual - Maintenance Windows Support Tool
## Prereq
- Administrators rights in MEMCM
- The MEMCM console must be installed on server or client

## HowTo
1. Select which collection you want to change/create/delete Maintenance Windows on
2. Click on button "Get existing Maintenance Windows" to get list of existing
3. Choose which months you will have Maintenance Windows for the selected collection
4. Select year
5. Select start time, hours and minutes
6. Select run-hours and run-minutes
7. Select how many days offset from patch-tuesday you want the Maintenance Windows to be
8. Select if the Maintenance Windows will apply to "Any", "SoftwareUpdatesOnly" or "TaskSequenceOnly"
9. If you are updating existing Windows you need to remove the old before by checking "Remove Old Maintenance Windows"
10. To update/create/remove click on button "Create Maintenance Windows"

A log file is created under c:\windows\logs with today's date where everything you do in the application is saved.
