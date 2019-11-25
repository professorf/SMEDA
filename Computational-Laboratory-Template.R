#
# Step 1. Insert your folder
#
direc="TrumpFearsYang" # Datasets must be in this folder
#direc="DemDebate3"
#direc="ClimateStrike"
files=dir(direc,"*.csv") # Datasets must be in .csv format
#
# Step 2. Pick a day to analyze
#
file=files[1]            # Choose a file to explore, e.g., 4 is Sekiro release day
#
# Read in the  file
#
df=read.csv(sprintf("%s/%s",direc,file), stringsAsFactors = F)
df$TweetDate=as.POSIXct(df$TweetDate,format="%a %b %d %H:%M:%S %z %Y",tz="GMT") # Convert a TweetDate into a usable format

if (F) { # T only gets English tweets
  EnRows=which(tolower(df$Language)=="en")
  df=df[EnRows,]
}

#
# Do a bar chart of the hourly volume to find spikes
#
h=hist(df$TweetDate,breaks="hours", main=sprintf("#%s/%s", direc, file),freq=T, xlab="HOUR", ylab="RETWEET #")

#
# Step 3. Pick an hour to analyze or an hour range
#
# Extract a dataframe of the spike or whatever hour you're interested in
# NOTE: The first hour is 0
#
#spikerows=which(as.integer(format(df$TweetDate,"%H"))==2) # Number (2) depends on visual inspection
spikerows=which(as.integer(format(df$TweetDate,"%H"))>=15 & as.integer(format(df$TweetDate,"%H"))<=15) # range example

spike=df[spikerows,]

#
# Find virals
#
tweetfreq=table(spike$Tweet)                   # Collapse Retweets
tweetsort=sort(tweetfreq,decreasing = TRUE)    # Sort tweets
totaltweets=sum(tweetsort)
totalunique=length(tweetsort)
totalnorts=length(which(tweetsort==1))
ttratio=totalunique/totaltweets
maxviral=tweetsort[1]
mtratio=maxviral/totaltweets



#
# At this point, tweetsort contains all the virals from most to least
# Example: tweetsort[1] is the most viral
# So you can start analyzing.
#
source("Computational-Laboratory-Functions.R") # load library
v=c(1:50)
t=as.character(names(tweetsort[v]))
dtm=DocumentTermMatrix(Corpus(VectorSource(t)))
mat=as.matrix(dtm)
d=dist(mat,method="binary")
h=hclust(d)
plot(h)

wordcloud(t,random.order=F,rot.per=0, col=brewer.pal(8, "Dark2"), min.freq=2)
#
# Possible things to try
#
# tweetsort[1]
# createDailyVolume(direc,files) # barchart of daily volume
 genSociogram(c(2,4:7))           # 8 is probably the most you should do

# getOverlap(1,2)                # Gets overvap between two groups
# see(1)                         # View tweetsort # in browser
# genBioHClust(100)
# cutoff(.5)                     # how many tweets until fraction reached
# tallyplot()                    # plot of virality
#

#
# Other: Math Modeling of Spike
#
# bar=createDailyVolume(direc,files)
# dfc=data.frame(x=1:length(files),y=bar) 
# plotSpike(dfc)
#

#
# Miscellaneous: Exploratory Data Code 
#

plotPareto()
N=cutoff(.25) # example: N=16
genBioHClust(1:N,2)
genBioHClust(1:N,10)
genWF(1)[1:2]
# ...
#genWF(16)[1:2]
#
# see(1)
# ...
# see(16)