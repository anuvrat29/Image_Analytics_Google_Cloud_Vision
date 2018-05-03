import tweepy,os,string
import xlwings as xw
import pandas as pd

ckey=""
csecret=""
auth = tweepy.AppAuthHandler(ckey,csecret)
api = tweepy.API(auth,wait_on_rate_limit=True,wait_on_rate_limit_notify=True)

def text_process(message):
    if message==[]:
        return ""
    else:
        no_punctuation = [char for char in message if char not in string.punctuation]
        no_punctuation = ', '.join(no_punctuation)
        return no_punctuation

def run_twitter():
    wb = xw.Book.caller()
    maxTweets = wb.sheets[0].range('E11').value
    if maxTweets is None:
        wb.sheets[0].range('B17').value = "Please Enter Number of tweets."
        return
    searchQuery = wb.sheets[0].range('E13').value
    if searchQuery is None:
        wb.sheets[0].range('B17').value = "Please Enter Keyword Query."
        return
    tweetsPerQuery = 100
    max_id = -1
    tweetCount = 0
    wb.sheets[0].range('B17').value = "Downloading maximum "+str(int(maxTweets))+" to "+str(int(maxTweets)+99)+" tweets."
    i=0
    imagesurl = []
    created,text,url,name,followers,following = [],[],[],[],[],[]
    loc,like,retweet,reply,hashtag,retweeted = [],[],[],[],[],[]

    while tweetCount < maxTweets:
        try:
            if (max_id <= 0):
                new_tweets = api.search(q=searchQuery, count=tweetsPerQuery)
            else:
                new_tweets = api.search(q=searchQuery, count=tweetsPerQuery, max_id=str(max_id-1))
        
            if not new_tweets:
                print("No more tweets found")
                break

            for tweet in new_tweets:
                if "media" in tweet._json["entities"]:
                    for image in  tweet._json["entities"]["media"]:
                        imagesurl.append(image["media_url"])
                        created.append(tweet._json["created_at"][:-11])
                        if "retweeted_status" in tweet._json:
                            text.append(tweet._json["retweeted_status"]["text"])
                        else:
                            text.append(tweet._json["text"])
                        name.append(tweet._json["user"]["name"].upper())
                        followers.append(tweet._json["user"]["followers_count"])
                        following.append(tweet._json["user"]["friends_count"])
                        if "RT" in tweet._json["text"].split(" ")[0]:
                            retweet.append(0)
                            retweeted.append("RETWEET")
                        else:
                            retweet.append(tweet._json["retweet_count"])
                            retweeted.append("")
                        like.append(tweet._json["favorite_count"])
                        url.append("https://twitter.com/"+str(tweet._json["user"]["id"])+"/statuses/"+tweet._json["id_str"])
                        if tweet._json["user"]["location"]=="":
                            loc.append("Unknown")
                        else:
                            loc.append(tweet._json["user"]["location"])
                        hashtags = [hashtag["text"] for hashtag in tweet._json["entities"]["hashtags"]]
                        hashtag.append(text_process(hashtags))
                        i=i+1

            tweetCount += len(new_tweets)
            wb.sheets[0].range('B17').value = "Scanned {0} tweets so far and found {1} images...".format(tweetCount,i)
            max_id = new_tweets[-1].id
    
        except tweepy.TweepError as e:
            print("some error : " + str(e))
            break

    wb.sheets[0].range('B17').value = "Scanning Completed"

    dataset_tweets = pd.DataFrame({"Image URLs":imagesurl,"DATE":created,"AUTHOR":name,"LOCATION":loc,"CONTENT":text,"FOLLOWERS":followers,"LIKES":like,"RETWEET":retweet,"ARTICLE_URL":url,"FOLLOWING":following,"HASHTAGS":hashtag,"POST_TYPE":retweeted},columns=["Image URLs","DATE","AUTHOR","LOCATION","CONTENT","ARTICLE_URL","LIKES","RETWEET","POST_TYPE","FOLLOWING","FOLLOWERS","HASHTAGS"])
    dataset_tweets = dataset_tweets.drop_duplicates("Image URLs", keep='first')
    unique_image = dataset_tweets.iloc[:,0].values
    dataset_stats = pd.DataFrame({"No of Tweets":[maxTweets],"Total Images":[i],"Unique Images":[unique_image.size]},columns=["No of Tweets","Total Images","Unique Images"])

    filepath = os.path.dirname(os.path.abspath(__file__))
    writer = pd.ExcelWriter(filepath+"/RealTimeUrl.xlsx")
    dataset_tweets.to_excel(writer,sheet_name="RealTimeUrls",index=False)
    dataset_stats.to_excel(writer,sheet_name="Stats",index=False)
    writer.save()
    wb.sheets[0].range('B17').value = "Congratulations...!!!  Collected "+str(i)+" images, out of those "+str(unique_image.size)+" is/are unique."