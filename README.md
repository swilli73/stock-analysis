# stock-analysis

## Overview of Project
The purpose of this analysis was not to analyze and create an output for the data, this was already done in the module. The purpose was to essentially take the "rough draft" code worked through during the lesson and refactor it to run quicker and more efficiently. This entails going through some of the same steps of writing the original code, but utilizing arrays and more variables to make it shorter and able to handle more data if need be.

## Results

### 2017 
![2017 Output](https://i.gyazo.com/5cf6bf27aeb0933511bfb5b341051d9c.png)

2017 appears to have been a very good year for growth, with every ticker appearing positive except for TERP. DQ, SEDG, ENPH, and FSLR are among the most successful, which would allow one to believe that they are the best to follow for the coming year. This is comparable to AY and RUN, which showed growth, however minimal.

![Original Run Time](https://i.gyazo.com/a285dbc4cf8402126e9ecaf0abe1e25e.png)
![Refactored Run Time](https://i.gyazo.com/76f758a8345ccb096f6eaf36f9e0ef33.png)

These are the script run times for 2017. While both short, the version with the refactored code runs much quicker. With a larger pool of data, the refactored code would fare much better processing times and save a bit of possible headache.

### 2018
![2018 Output](https://i.gyazo.com/7a66c61785ed428f68b460f807da0fe3.png)

2018 negatively exceeded expectations by throwing off the stats a bit unpredictably. AY, SEDG, TERP, and VSLR show comparatively smaller losses, but the rest took a pretty big hit, especially DQ and JKS as they are among the highest. RUN and ENPH are the only tickers that continue to grow, both experiencing quite a large return this year. While RUN also grew last year, it was a small amount, so I would definitely place my bets on ENPH's continued success.

![Original Run Time](https://i.gyazo.com/c342b0539d27b5faca5a96ce1e4690a5.png)
![Refactored Run Time](https://i.gyazo.com/40c7b775fbdc47f1abe9f326a1a10023.png)

As with the 2017 example, these are the script run times for 2018. The results are much faster for the refactored code, and overall the quickest for both years. 

### Differences in Code
![Original Code](https://i.gyazo.com/6c8e345c325bc67daa6461a4ca5fc136.png)
![Updated Code](https://i.gyazo.com/955e04b8db6555bd54cb81a33e82bb62.png)
In my opinion, one of the largest differences in the code, and why one runs faster than the other, is near the end of the script with the table output. The refactored code utilizes an array for the values, while the module code goes through the work of establishing "ticker", "totalVolume", "startingPrice", and "endingPrice" as unrelated entities. From my knowledge, this seems to increase the work the program needs to do compared to utilizing arrays and variables. 

![Array](https://i.gyazo.com/35b65d9fbf999da6ad6d3ab8e7c44607.png)
Having all of your output variables related to each other is not only cleaner and more organized, but the computer does seem to handle it a lot better. After having worked through the material this is the largest thing that seems to stand out to me. 

## Summary

### What are the advantages or disadvantages of refactoring code?
  - The advantages are shorter processing time, code that is cleaner and easier to read, as well as sharpening coding skills by seeking more efficient solutions.
  - The disadvantages lie in time and lack of knowledge. Going into this blind, I would have had little idea on how to refactor this code without the starter code provided. I still had to do a lot of research and code comparison (though a lot of it was also reusing code from the module) and could not imagine this task on a larger scale. I'm thankful that I now have this refactored code as a template for future VBA scripts as it is very efficient and powerful, I can believe that it would be a good resource. The time it takes to refactor code is less of a factor when you have something to compare it to, and in my opinion, writing original, rougher code still takes much longer and is more work. 

### How do these pros and cons apply to refactoring the original VBA script?
  - While more experienced now, I still have a lot to learn with VBA and am thankful for this project. I am more hesitant to it than I am Python (just speaking from a lens of being on Week 3), but I can also see and appreciate the merits of how I could reuse code from the original to refactor that same code and make it "better". I did go into a bit of detail in the previous question, but I believe that being candid with the process encourages growth and retention of these new skills rather than copying something and moving on for a grade. The advantages heavily outweigh the disadvantages, and this is efficient code that can handle even more tickers and stock data if need be. The original and refactored code still share a lot, but the refactored code works a lot smarter, while the original works harder and drags out establishing the output variables.
