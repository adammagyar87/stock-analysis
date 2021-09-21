# Green Stock Analysis
## Overview:
#### A list of green stocks was compiled for the client to analyze how they performed in 2017 and 2018. We utilized VBA to quickly present the data from the start to end of both years. After the initial build, we refactored the code in VBA to run more efficiently if given a larger dataset.

## Results:
#### In 2017 DQ was a very strong performer with just under a 2x price increase. However almost all of that progress was lost in 2018. The starting price for DQ in 2017 was $19.85 and it ended 2018 at $$23.40.
#### Most of the Green Stocks in the list had a strong 2017 followed buy a big downturn in 2018. The tables below highlight this drastic turn.

![image](https://user-images.githubusercontent.com/87042597/134109470-c8ed1f60-92ec-48e3-9356-6c808489711f.png) ![image](https://user-images.githubusercontent.com/87042597/134109410-39d469c7-965b-4db0-b961-16b91163e80c.png)

#### However there were two stocks in this list which had positive gains in both years, "ENPH" and "RUN" with "ENPH" being a big gainer in both.


## Summary:
#### Refactoring the code did show an improvement on the runtime compared to the original version. The run times on my workstation decreased from roughly 0.79 seconds to under 0.73 seconds, as shown in the screenshots below.

![image](https://user-images.githubusercontent.com/87042597/134109993-c299b6b7-b53e-4c5e-94d2-76810c5fd38f.png)

#### A potential disadvantage of refactoring the code would be if it didn't include proper comments and it was difficult to understand what certain snippets of code were doing. This could cause a lot of confusion and might potentially make the process more trouble than it was worth. 

#### An advantage of the original code we wrote was that it less complicated than the refactored version. We only set up a single array and it was well defined in our code. The refactored version used code to auto-populate the output arrays. Since we are doing this for a friend and his parents, its reasonable to assume they are not very experienced with VBA. If this is the case it would seem simplicity would be valued over efficiency.

#### The refactored code did however run faster than the original version. If the person this was built for would like to expand the program to many more stock tickers, then this efficiency would provide a better end user experience.
