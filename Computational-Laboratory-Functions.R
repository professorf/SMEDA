################################################################################
# Generate social network (sociograms) for virals
################################################################################
if (require("tm")==F) {
  install.pacakges("tm")
  library(tm)
}
if (require("wordcloud")==F) {
  install.packages("wordcloud")
  library(wordcloud)
}
if (require("igraph")==F) {
  install.packages("igraph")
  library(igraph)
}

genSociogram=function(vrange=c(1:2)) {
  # Create the edges for the mention network
  edge_list=c()
  for (v in vrange) {
    vrows=which(spike$Tweet==names(tweetsort[v]))
    for (i in 1:length(vrows)) {
      from_edge=spike$UserName[vrows[i]]
      to_edges=unlist(strsplit(spike$Mentions[vrows[i]],","))
      for (n in to_edges) {
        #      edge=c(from_edge,n,names(tweetsort[v]),v) # from, to, label, group # TO SAVE MEMORY
        edge=c(from_edge,n) # from, to, label, group
        edge_list=rbind(edge_list,edge)
      }
    }
  }
  
  # Create a data frame
  #colnames(edge_list)=c("from","to", "label", "group")
  colnames(edge_list)=c("from","to")
  #write.csv(edge_list,"edges.csv") # Write out edges for other programs
  
  #dfedges=data.frame(edge_list) # Do not create labels & groups for huge networks
  dfedges=data.frame(from=edge_list[,"from"],to=edge_list[,"to"]) # TO SAVE MEMORY
  
  # Create the graph object
  g=graph_from_data_frame(dfedges)
  
  # Group by edges (REDO: too slow)
  # com=cluster_edge_betweenness(as.undirected(g))
  edges=c()    # All edges
  group=list() # List of verices
  for (v in vrange) {
    vrows=which(spike$Tweet==names(tweetsort[v]))
    vedges=c()  # Edges for just the current viral in vrange
    for (i in 1:length(vrows)) {
      from_edge=spike$UserName[vrows[i]]
      to_edges=unlist(strsplit(spike$Mentions[vrows[i]],","))
      for (n in to_edges) {
        edge=c(from_edge,n) # from, to, label, group
        edges=rbind(edges,edge)
        vedges=rbind(vedges,edge)
      }
    }
    verts=unique(paste(vedges)) # All the unique vertices in a viral group
    group=c(group,list(verts))  # A list of all the ~unique verts in all viral groups
  }
  
  # Set vertext attributes & draw sociogram
  #V(g)$color=com$membership+1
  for (i in 1:length(V(g)$name)) {
    for (j in 1:length(group)) {
      if (V(g)$name[i] %in% group[[j]]) { # Color is first group found
        V(g)$color[i]=j
        break
      }
    }
  }
  
  V(g)$size=ifelse(strength(g)>10,15,5) # Hack: Fix node sizes
  # V(g)$size=log(strength(g))
  # V(g)$size=strength(g)
  if (T) {                              # Change to T to plot sociogram
    set.seed(777)                         # Seed necessary otherwise random network
    plot(g,layout=layout_with_kk (g),vertex.label=NA,edge.label=NA, edge.arrow.size=.5)
    if (F) {
      ug=as.undirected(simplify(g))
      gr=cluster_fast_greedy(ug)
      # dendPlot(gr)
      plot(gr,g, layout=layout_with_kk (g),vertex.label=NA,edge.label=NA, edge.arrow.size=.5)
    }
    legend("right",legend=as.character(vrange),bty="y", pch=19, cex=1.25,col=categorical_pal(8))
    title(sprintf("Sociogram: Virals %s", paste(vrange,collapse=",")))
    #V(g)$color=gr$membership
    #tkplot(g,layout=layout_with_kk (g),vertex.label=NA,edge.label=NA, edge.arrow.size=.5)
  }
}

getOverlap=function(grpA=1,grpB=2) {
  groupA=grpA
  groupB=grpB
  tweetA=names(tweetsort[groupA])
  tweetB=names(tweetsort[groupB])
  tweetArows=which(spike$Tweet==tweetA)
  tweetBrows=which(spike$Tweet==tweetB)
  A=c()
  for (i in tweetArows) {
    A=c(A,spike$UserName[i])
    mentions=unlist(strsplit(spike$Mentions[i],","))
    A=c(A,mentions)
  }
  B=c()
  for (j in tweetBrows) {
    B=c(B,spike$UserName[j])
    mentions=unlist(strsplit(spike$Mentions[j],","))
    B=c(B,mentions)
  }
  A=unique(A)
  B=unique(B)
  tweeners=intersect(A,B)
  tweeners
}

#
# Given a directory (assumes all .csv files)
# Returns a vector of daily volumes, which you can plot
#
createDailyVolume=function(dir, fs) {
  bar=c()
  for (f in fs) {
    fn=sprintf("%s/%s", dir, f)
    print(sprintf("Reading: %s...",fn))
    df=read.csv(fn)
    bar=c(bar, length(df$TweetID))
  }
  bar
}
#
# Plots an hourly bar chart
#
plotHourlyVolume=function(df) {
  h=hist(df$TweetDate,breaks="hours", main=sprintf("#%s/%s", direc, file),freq=T, xlab="HOUR", ylab="RETWEET #")
  h
}

see=function (viral,i=1) {
  tweet=tweetsort[viral] # Tweet associated with index
  vrows=which(spike$Tweet==names(tweetsort[viral])) # All users retweeting tweet
  browseURL(paste("https://twitter.com/",spike$UserName[vrows[i]],"/status/",spike$TweetID[vrows[i]],sep="")) # Vary index for different users
}


modelSpike=function(curve, b=2,shift=2,tall=100000) # Typically need to modify just tall
{
  fit=nls(y~tall/10^((log(x/shift)/log(b))^2), curve,
          start=list(shift=shift,tall=tall,b=b),trace=T)->fit
  fit
}

plotSpike=function(dfc) {
  fit=modelSpike(dfc,tall=max(dfc$y))
  s=summary(fit)
  bx=barplot(dfc$y,ylim=c(0,max(dfc$y)+max(dfc$y)/10))
  curve(s$coefficients["tall",1] / 10^((log(x/s$coefficients["shift",1])/log(s$coefficients["b",1]))^2), add=T)
  title(sprintf("%.2f/10^((log(x/%.2f)/log(%.2f))^2",s$coefficients["tall",1],s$coefficients["shift",1],s$coefficients["b",1]))
  text(bx,dfc$y+max(dfc$y)/10/2,dfc$y, las=1,cex=0.75)
}

cutoff=function(level) { # Level entered as a decimal
  total=sum(tweetsort)
  tally=0
  for (i in 1:length(tweetsort)) {
    tally=tally+tweetsort[i]
    if ((tally/total)>=level) break;
  }
  i
}

tallyplot=function() {
  x=c()
  y=c()
  total=sum(tweetsort)
  tally=0
  for (i in 1:length(tweetsort)) {
    tally=tally+tweetsort[i]
    x=c(x, i)
    y=c(y,tally/total)
  }
  plot(x,y)
}

tallyplotabs=function() { # Absolute Tally values, for shape purposes
  x=c()
  y=c()
  total=sum(tweetsort)
  tally=0
  for (i in 1:length(tweetsort)) {
    tally=tally+tweetsort[i]
    x=c(x, i)
    y=c(y,tally) #/total)
  }
  plot(x,y,xlim=c(0,total), ylim=c(0,total))
  lines(c(0,totaltweets),c(0,totaltweets))
  onelen=length(which(tweetsort==1))
  lines(c(totalunique-onelen,totalunique-onelen),c(0,totaltweets))
  lines(c(totalunique,totalunique),c(0,totaltweets))
  title("Pareto Chart")
  xpos=totalunique-onelen
  text(xpos,totaltweets,as.character(xpos))
  text(totalunique,totaltweets,as.character(totalunique), pos=4)
}

plotPareto=function() { # Absolute Tally values, for shape purposes
  x=c()
  y=c()
  total=sum(tweetsort)
  tally=0
  for (i in 1:length(tweetsort)) {
    tally=tally+tweetsort[i]
    x=c(x, i)
    y=c(y,tally) #/total)
  }
  plot(x,y,xlim=c(0,total), ylim=c(0,total))
  lines(c(0,totaltweets),c(0,totaltweets))
  onelen=length(which(tweetsort==1))
  lines(c(totalunique-onelen,totalunique-onelen),c(0,totaltweets))
  lines(c(totalunique,totalunique),c(0,totaltweets))
  title("Viral Pareto Diagram")
  xpos=totalunique-onelen
  text(xpos,totaltweets,as.character(xpos))
  text(totalunique,totaltweets,as.character(totalunique), pos=4)
}

tallyplotpct=function() {
  x=c()
  y=c()
  total=sum(tweetsort)
  lentw=length(tweetsort)
  tally=0
  for (i in 1:lentw) {
    tally=tally+tweetsort[i]
    x=c(x, i/total)
    y=c(y,tally/total)
  }
  plot(x,y,xlim=c(0,1),ylim=c(0,1))
}

#
# plotAntiPareto: (experimental)
#
#
# Calculate the anti-pareto
#
plotAntiPareto = function () {
  invtotal=c()
  for (i in 1:25) invtotal=c(invtotal,length(which(tweetsort==i)))
  barplot(invtotal[2:25])
  
  cutoff50=cutoff(.5)
  cutoff25=cutoff(.25)
  barplot(tweetsort[1:25])                       # Plot, maybe do scree test
  plot(2:200,tweetsort[2:200])
}

#
# genHClust: Does a user bio similarity for a tweet. requires $bio field
#
# entry: tweetsort
# exit : rows
#
# do indirect addressing to see a particular bio: spike$Bio[rows[#]] to see a bio
genHClust=function(virnum) {
  rows=which(spike$Tweet==names(tweetsort[virnum]))
  text=spike$Bio[rows]
  dtm=DocumentTermMatrix(Corpus(VectorSource(text)))
  mat=as.matrix(dtm)
  d=dist(mat,method="binary")
  h=hclust(d)
  plot(h, xlab="User Who Retweeted", ylab="Bio Difference", main="Bio Differences")
  rows
}
#
# genWC: Generates a bio wordcloud for a given tweet.
#
genWC=function(sortnum) {
  rows=which(spike$Tweet==names(tweetsort[sortnum]))
  text=spike$Bio[rows]
  wordcloud(text,random.order=F,rot.per=0,col=brewer.pal(8,"Dark2"),max.words=250)
}

#
# genWF: Generates a bio frequency table for a given tweet
#
genWF=function(sortnum, removesw=T) {
  rows=which(spike$Tweet==names(tweetsort[sortnum]))
  text=tolower(spike$Bio[rows])
  comb=unlist(strsplit(text, " "))
  if (removesw) comb=removeWords(comb,stopwords())   # remove with punctuation
  comb2=gsub("[^a-zA-Z0-9]","", comb)
  if (removesw) comb2=removeWords(comb2,stopwords()) # try again
  comb3=comb2[which(comb2!="")]                      # remove empty words
  ttext=table(comb3)
  sort(ttext,decreasing=T)                           # sort
}

#
# getWordOverlapTable
# Entry: 
#   u: the user vector that you are interested in comparing bios
#   rowvec: The row vector returned by genHClust
#
# This function uses the row vector returned by genHClust -> rowvec
#
#
# The idea is that if the length of the row vector (rowvec) was 3
# The maximum frequency would be 3 and that would indicate that the word
# was common across all three users
#
# A clever way of determining overlap.
#
getWordOverlapTable=function(u,rowvec) { # not sure if this is correct
  wordlist=c()
  for (row in u) {
    bio=spike$Bio[rowvec[row]]
    sbio=tolower(unlist(strsplit(bio, " ")))
    cbio=gsub("[^a-z]","",sbio)
    cbio=unique(cbio[which(cbio!="")]) # Get rid of blanks & make unique
    print(cbio)
    wordlist=c(wordlist,cbio)
  }
  twordlist=sort(table(wordlist), decreasing=T)
  twordlist
}


#
# Bioprint functions
#

#
# BioTallyPlot: 
#   Given a viral tweet, this function first takes all the bios of all the users
#   who retweeted the tweet, and creates a table of all the unique words and
#   a tally for each word (how many times the word occurs)
#   It then sorts all the words in decreasing order.
#   Then for each word, in order, it sums up all the prevous word tallies and 
#   divdes by the total # of words to get a % of tweets accounted for.
#
#
BioTallyPlot=function(sortnum, removesw=T) {
  sortfreq=genWF(sortnum,removesw)
  x=c()
  y=c()
  total=sum(sortfreq)
  tally=0
  for (i in 1:length(sortfreq)) {
    tally=tally+sortfreq[i]
    x=c(x, i)
    y=c(y,tally/total)
  }
  plot(x,y, xlab="Word #", ylab="% of Total Volume")
  title(sprintf("Cumulative Distribution for Viral Tweet # %d",sortnum))
}

BioCutoff=function(cutoff, sortnum, removesw=T) {
  sortfreq=genWF(sortnum,removesw)
  total=sum(sortfreq)
  len=length(sortfreq)
  subtot=0
  for (i in 1:len) {
    subtot=subtot+sortfreq[i]
    pct=subtot/total
    if (pct>cutoff) break
  }
  i
}

genBioHClust=function(subset, n=0) {
  vb=c()
  for (i in subset) {
    b=genWF(i)
    nb=names(b)
    if (n!=0) nb=names(b[1:n])
    pb=paste(nb, collapse=" ")
    vb=c(vb,pb)
    #    print(vb)
  }
  dtm=DocumentTermMatrix(Corpus(VectorSource(vb)))
  mat=as.matrix(dtm)
  d=dist(mat,method="binary")
  h=hclust(d)
  plot(h)
}

getWords=function(v) {
  words=tolower(unlist(strsplit(names(tweetsort[v]), " ")))
  words=gsub("[^a-zA-Z]","", words)
  NonBlankRows=which(words!="")
  words=words[NonBlankRows]
  words
}

