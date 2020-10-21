# Google Spreadsheet Automation 

### A bit about me... 
Hello! I'm a machine learning engineer at Civis Analytics where I 
1. Developed/ship our house forecasting model to HMP 
2. Created a resource allocation service/dashboard for SMP and the Biden campaign 
3. Developed/manage our flight radar ad testing system for our senate clients
3. Manage modeling-as-a-service products for our clients

### What is this repo? 
In progressive politics its important to get out numbers and get eyes on them _quickly_. Autoupdating spreadsheets are a convenient way to track important tables in an accessible way + maintain version history. This repo is meant as a simple boilerplate to set you up to do this fast. 

![](https://giphy.com/gifs/l4RKhOL0xiBdbgglFi)

### Steps ### 
0. Set up your gcp (won't cover this)
1. Create a service account in your gcp 
    a. Go to https://console.cloud.google.com/
    b. Search service accounts 
    c. + Create service account and name it 
    d. Click on your created service account and Add Key 
    e. Create New Key, Json, and save it in your bash_profile 
    with format 

    export SA_PASSWORD=`cat << EOF <paste json> EOF`
2. Create a spreadsheet by setting `CREATE_NEW_SPREADSHEET=True` in `main.py` and write df to sheet
4. Set `CREATE_NEW_SPREADSHEET=False` to keep your old sheet and choose conditional formatting 

### Questions? 
@ashe_cs <br/>
ashe.magalhaes@gmail.com 
