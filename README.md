# Google Spreadsheet Automation 

### A bit about me... 
Hello! I'm a machine learning engineer at Civis Analytics where I 
1. Developed/ship our house forecasting model to HMP <br/>
2. Created a resource allocation service/dashboard for SMP and the Biden campaign <br/>
3. Developed/manage our flight radar ad testing system for our senate clients <br/>
3. Manage modeling-as-a-service products for our clients <br/>

### What is this repo? 
In progressive politics its important to get out numbers and get eyes on them _quickly_. Autoupdating spreadsheets are a convenient way to track important tables in an accessible way + maintain version history. This repo is meant as a simple boilerplate to set you up to do this fast. <br/>

![](https://media.giphy.com/media/l4RKhOL0xiBdbgglFi/giphy.gif)

### Steps ### 
0. Set up your gcp (won't cover this)
1. Create a service account in your gcp  <br/>
    a. Go to https://console.cloud.google.com/ <br/>
    b. Search service accounts <br/>
    c. + Create service account and name it <br/>
    d. Click on your created service account and Add Key  <br/>
    e. Create New Key, Json, and save it in your bash_profile 
    with format <br/>

    export SA_PASSWORD=`cat << EOF <paste json> EOF`<br/>
2. Create a spreadsheet by setting `CREATE_NEW_SPREADSHEET=True` in `main.py` and write df to sheet<br/>
4. Set `CREATE_NEW_SPREADSHEET=False` to keep your old sheet and choose conditional formatting <br/>

### Questions? 
@ashe_cs <br/>
ashe.magalhaes@gmail.com 
