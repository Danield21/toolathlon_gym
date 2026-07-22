You are a market research analyst responsible for compiling a weekly market news digest for your investment newsletter. Your workspace contains Newsletter_Template.md with the email template format and Subscriber_List.txt with the distribution list address.

Begin by fetching recent financial news for all five tracked stocks: GOOGL, AMZN, JPM, JNJ, and XOM. Pull the latest news headlines for each ticker symbol from the financial market data service.

Then capture this digest in our team knowledge base by creating a database titled "Market News Digest" with the following properties: Title (title type), Symbol (rich_text type), Publisher (rich_text type), Published_Date (rich_text type), and Summary (rich_text type). Add at least five entries to this database, one for each news article. For each entry, truncate the summary or description to at most 200 characters.

Next, prepare an investor sentiment poll using an online survey form titled "Investor Market Sentiment Survey" with exactly four questions. The first question must be "How would you rate the overall market outlook for the next 30 days?" with multiple choice options: Very Bullish, Bullish, Neutral, Bearish, Very Bearish. The second question must be "Which sector do you expect to outperform in the next quarter?" with multiple choice options: Communication Services, Consumer Cyclical, Financial Services, Healthcare, Energy. The third question must be "What is your primary investment concern right now?" with multiple choice options: Inflation, Interest Rates, Geopolitical Risk, Earnings Slowdown, Valuations. The fourth question must be "How likely are you to increase your equity allocation in the next month?" with multiple choice options: Very Likely, Likely, Neutral, Unlikely, Very Unlikely.

Finally, send an email from newsletter@fund.example.com to subscribers@newsletter.example.com with a subject that contains all of the words "Weekly", "Market", and "Digest". The body should include highlights from the news you gathered, covering at least three of the five tracked stocks.

When everything is in place, call claim_done.
