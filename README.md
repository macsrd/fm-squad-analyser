# fm-squad-analyser
Script made to help to find more suitable players for my team in Football Manager 2021. The idea is to use exported file from game menu and calculate player rating for each positon using Python pandas library.

In game I use to play English Lower League Team - Concord Rangers. To save money in the club I use to do scouting on my own by inviting multiple players to club on tests and letting my staff to 'uncover' player skills. 

First version of script is made to calculate my current players rating and suitability for each position, to do this I use skills buckets saved in the Excel file, each file is getting imported then to dataframe as list and basing on these lists ratings for each positons are being calculated. Currently for each positions are three buckets, first bucket (_1) of most important skills that is being multiplied by 1, second bucket (_2) by 0.8 and third bucket (_3) by 0.6.
