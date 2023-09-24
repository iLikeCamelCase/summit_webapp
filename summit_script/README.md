# paystub_script

paystub_script is a python script which reads data from Summit Reforestation Planter paystubs and displays it on an excel spreadsheet

## Usage

Populate 'paystubs' folder with your planter paystubs, run paystub_script.

An excel file will be created.



## To-do

PROGRAMMER FEATURES:
- [x] change x-axis from 1,2,3...50 to dates
- [ ] proper logging
- [x] implement regex method for parsing data

USER FEATURES:
- [ ] ability to choose trees or pay, one or other, bargraph only

Tree/Pay Charts & Features:
- [ ] function for creating tree and pay charts
- [ ] ability to plot average (trees or pay)

Side Data:
- [x] data by contract
- [x] average working day treecount
- [x] average working day gross pay
- [x] total tree count
- [x] total pay
- [ ] breakdown by centage (pie chart thingy)
- [ ] average centage
- [x] total num contracts
- [x] total num blocks
- [x] best day
- [x] worst day

## Bug fixes

- [x] populate days off list of days
- [x] may 31st not a thing?
- [x] regex only matches first instance
- [ ] centage data