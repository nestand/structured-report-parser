# structured-report-parser

A Python script that updates a formatted Excel report template with values from a search result file.

It matches rows based on the **report name**, copies over selected fields, and writes them back to a new Excel file **without breaking formatting**.  
Originally developed for a client delivery — now refactored and shared in a neutral form to improve structure, flexibility, and maintainability.

---

## 🚀 Features

- ✅ Matches rows by "Report Name"
- ✅ Updates mapped fields (like report type, time taken, volume)
- ✅ Preserves Excel formatting using `openpyxl`
- ✅ Saves to a new file with a timestamp
- ✅ Refactored into functions for clarity
- 🔜 Looking to improve structure 

---

## 🧠 Why I Shared This

> The project worked fine in delivery, but I’m now looking for **developer feedback**.  
> I’d love to hear from more experienced Python devs on:
> 
> - Structuring this better (classes? CLI? library?)
> - Making it more flexible / reusable
> - Adding validation, testing, or even web input? OR very welcome some other aspects/ideas too ;) 

So Feel free to drop feedback, issues, or ideas 🙏

Thanks!
---

## 🛠 Tech Stack

- Python 3.11
- Conda environment (see below)
- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)



