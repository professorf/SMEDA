if (!require("rtweet")) install.packages("rtweet")
library(rtweet)
#
# ANALYTICS TEAM: FILL THIS IN FIRST
#
create_token(
  app = "",
  consumer_key = "",
  consumer_secret = "",
  access_token = "",
  access_secret = ""
)
# See reading:
# http://professorf.com/wp-content/uploads/2016/10/How-To-Register-As-A-Developer-And-Get-Your-Keys.pdf

#
# ANALYTICS TEAM: Modify one of the following and comment out the other
#
#df=search_tweets(q="from:@prattprattpratt",n=250000,include_rts=T, retryonratelimit=F) # Example: searching for people
#df=search_tweets(q="#roswellNM",n=100,include_rts=T, retryonratelimit=F) # Example:  searching for hashtags
df=search_tweets(q="#MondayMorning", n=1000, since="2019-11-25",until="2019-11-26", include_rts=T, retryonratelimit=T) # Example:  searching for hashtags
#
# Get all the tweets, then just the original tweets
#
NotRTsNorAts = which(df$is_retweet==F & grepl("^@", df$text)==F)
odf=df[NotRTsNorAts,] # Original Tweets that are not RTs nor @replies

#
# Summary Stats
#

total=length(df$text) # All tweets including RTs & @mentions
ototal = length(odf$text)
avg=mean(odf$retweet_count)
sdv=sd(odf$retweet_count)
mx=max(odf$retweet_count)
mn=min(odf$retweet_count)
qt=quantile(odf$retweet_count)
medin=median(odf$retweet_count)
followers=df$followers_count[length(odf$followers_count)] # Get the last follower count
viral.cutoff=avg+2*sdv
print(sprintf("Followers: %d", followers))
print(sprintf("Total tweets: %d", total))
print(sprintf("Total original tweets: %d", ototal))
print(sprintf("Max retweet: %d", mx))
print(sprintf("Min retweet: %d", mn))
print(sprintf("Average retweet: %0.2f", avg))
print(sprintf("Average Retweet/Followers: %0.4f", avg/followers*100))
print(sprintf("SDev retweet: %0.2f", sdv))
print(sprintf("Viral Cutoff: %0.2f", viral.cutoff))
print(qt)
print(sprintf("Median: %0.2f", medin))
print(sprintf("3rd Quartile: %0.2f", qt[4]))

save_as_csv(df,"rtweet.csv")
