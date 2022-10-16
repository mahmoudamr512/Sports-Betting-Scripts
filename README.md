# Sports Betting Website 

## Overview: 

My Client "Ryan" is a client interested in Sports Betting. He is interested to gather data from multiple games and multiple websites in different sports (MLB & NFL for now).

To help him have a way to determine quickly which games/players are worth betting on, I have created a series of scraping scripts for him. 

Those scripts are running on Selenium (Manual Run) to collect specific data from the websites and append to his "XLS" file in his documents. 

## Changes to Scripts to work on any PC:

If you're interested to make this work on your PC, you have to change each script's path. 

Also, if you're on Windows, sometimes OpenPyXL can do some problems bey creating an empty-byted XLS file (For this problem, you can always ensure to save a workbook before start writing to it).

If you're on UNIX, then you won't find this problem.

<smalL><strong>PS: When making those scripts, they are meant to run separately, each script on his own and can be copied on any PC (Independelty, not as package). Therefore, each script is considered to be self-sufficient. That's why, you won't find a universal export function that is there once and imported dunamically.</strong></small>

## Finished Scripts & Websites:


* [X] MLB-Betonline
* [X] MLB-DraftKings
* [X] MLB-Fanduel
* [X] MLB-FoxBet
* [X] NFL-Betonline
* [X] NFL-DraftKings
* [X] NFL-BetMGM
* [X] NFL-Barsatool
* [X] NFL-Ceasers
* [X] NFL-Faundel
* [X] NFL-FoxBet