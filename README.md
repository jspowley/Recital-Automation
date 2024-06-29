# Recital Scheduling Automation

### The impact: 

Reduction of scheduling times by about 20 hours, 2-4 times a year. Due to the small scale of operations and limited human capital, this means turnaround times on event programming have reduced by about a month, resulting in better organized events, and happier parents and teachers.

### Brief Summary:

Sometimes simple/short code can yield significant improvements to a business, especially where technical skillsets have been less avaiable earlier in a businesses operations. This is an example of a fix to a local music studios recital registration and scheduling process. Currently, their CRM provides a form output impractical to the average Excel user. The user instead can drop the Input Excel file into the script folder, run the script via shortcut, and have the preprocessed data loaded directly into a scheduling template. Future iterations will then generate schedules for sound techs and recital programs for the audience automatically, through Latex and through Canva mailmerge.

### Key Design Points

- Ideally the MS office suite would be avoided entirely, by integrating the data import directly to the website and linking the file output directly to an online application such as google sheets. The business owner has declined these features for cybersecurity reasons, and as such, minimmal manual configuration is required to run the process in its current state.
- Taking a look at the Input.xlsx file, we see an entire able is contained in a single cell, for each teachers submissions. We use look-ahead and look-behind RegEx, combined with a lazy search, to isolate individual elements from the table. This requires specific cofiguration when changes are made to the form.
- Configuration may change depending on the R Executable's location.
- Flexibility is maximized by using RLang to assign columns flexibily, allowing form fields to all be configured in one place. The only exceptions to this are operations which seek to merge/add columns together, which need to be specificalluy indexed or defined.

This project will continue to grow as the studio requests more changes. No personal information has been revealed in this project, an permission to post the materials conatined herein have been granted by the studio owner.
