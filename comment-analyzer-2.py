import os
from re import X
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Alignment, Font
from googleapiclient.discovery import build
from transformers import pipeline
import praw
from datetime import datetime, date, timedelta
# import tweepy


# NOTES FOR BEFORE RUNNING THE PROGRAM
# To avoid exceeding API requests to Youtube
# There is test data instead of real-time API data
# When running the program for real:
# Remove commented code under titles: "*** ACTUAL DATA FROM YOUTUBE API ***"
# And comment out code under titles: "*** FAKE DATA FOR TESTING PURPOSES ***"


# CREDENTIALS (in .env file, and ignored by git with .gitignore file):
load_dotenv()
yt_api_key = os.getenv("youtube_api_key")
# reddit_client_id = os.getenv("reddit_client_id")
# reddit_client_secret = os.getenv("reddit_client_secret")
# reddit_username = os.getenv("reddit_username")
# reddit_password = os.getenv("reddit_password")
# twitter_api_key = os.getenv("twitter_api_key")
# twitter_api_key_secret = os.getenv("twitter_api_key_secret")
# twitter_bearer_token = os.getenv("twitter_bearer_token")
# twitter_access_token = os.getenv("twitter_access_token")
# twitter_access_token_secret = os.getenv("twitter_access_token_secret")


def main():
    print("This may take 1-10 minutes (takes 3 seconds per comment)")
    wb = load_workbook("Comment Analyzer.xlsx")
    ws = wb.active
    youtube = build("youtube", "v3", developerKey=yt_api_key)
    video_data = get_episode_data(ws, youtube)
    video_title = video_data[0]
    yt_video_id = video_data[1]
    yt_video_date = video_data[2]
    episode_number = video_data[3]
    podcast_guest = video_title.partition(":")[0]
    prep_comment_ws(ws, wb)
    get_yt_comments(youtube, yt_video_id, ws, token="", x=0, row_number=2)
    # get_reddit_comments(video_title, ws)
    # get_twitter_comments(video_title, yt_video_date, ws)
    totals = sort_comments_by_sentiment(wb, ws)
    total_positive = totals[0]
    total_neutral = totals[1]
    total_negative = totals[2]
    wb.save(filename="Comment Analyzer - Complete.xlsx")
    print("Podcast Episode #", episode_number)
    print("Podcast Guest: ", podcast_guest)
    print("\n")
    total_comments = total_positive + total_neutral + total_negative
    percent_positive = str(total_positive * 100 / total_comments) + "%"
    percent_neutral = str(total_neutral * 100 / total_comments) + "%"
    percent_negative = str(total_negative * 100 / total_comments) + "%"
    print("Total Positive Comments: ", total_positive, " (", percent_positive, ")")
    print("Total Neutral Comments: ", total_neutral, " (", percent_neutral, ")")
    print("Total Negative Comments: ", total_negative, " (", percent_negative, ")")


def get_episode_data(ws, youtube):
    episode_number = ws["B1"].value
    # If no podcast episode is specified, use the latest episode
    if episode_number == None:
        latest_episode_data = get_latest_episode(youtube)
        episode_number = str(latest_episode_data[0])
        yt_video_id = latest_episode_data[1]
        video_title = latest_episode_data[2]
        yt_video_date = latest_episode_data[3]
    else:
        # # *** ACTUAL DATA FROM YOUTUBE API ***
        search_response = (
            youtube.search()
            .list(
                part="snippet",
                maxResults=1,
                q="Lex Fridman Podcast #" + str(episode_number),
            )
            .execute()
        )
        yt_video_id = search_response["items"][0]["id"]["videoId"]
        video_response = youtube.videos().list(part="snippet", id=yt_video_id).execute()
        video_title = video_response["items"][0]["snippet"]["title"]
        yt_video_date = search_response["items"][0]["snippet"]["publishedAt"][0:10]
        # *** END OF ACTUAL DATA ***
        # *** FAKE DATA FOR TESTING PURPOSES ***
        # yt_video_id = "ZFntEFXKDHM"
        # video_title = "Kate Darling: Social Robots, Ethics, Privacy and the Future of MIT | Lex Fridman Podcast #329"
        # episode_number = "329"
        # yt_video_date = "2022-10-15"
        # *** END OF FAKE DATA ***
    return (video_title, yt_video_id, yt_video_date, episode_number)


def get_latest_episode(youtube):
    # *** ACTUAL DATA FROM YOUTUBE API ***
    lex_fridman_yt_channel_id = "UCSHZKyawb77ixDdsGog4iWA"
    search_response = (
        youtube.search()
        .list(
            part="id, snippet",
            type="video",
            order="date",
            channelId=lex_fridman_yt_channel_id,
            maxResults=5,
        )
        .execute()
    )
    for x in range(4):
        yt_video_id = search_response["items"][x]["id"]["videoId"]
        video_response = youtube.videos().list(part="snippet", id=yt_video_id).execute()
        video_title = video_response["items"][0]["snippet"]["title"]
        yt_video_date = str(search_response["items"][x]["snippet"]["publishedAt"])[0:10]
        if "Lex Fridman Podcast" in video_title:
            location_of_hash = video_title.find("#")
            episode_number = ""
            for x in range(1, 5):
                if (location_of_hash + x) < len(video_title):
                    if video_title[location_of_hash + x].isnumeric():
                        number_to_add = video_title[location_of_hash + x]
                        episode_number += number_to_add
            episode_number = episode_number
            break
    # # *** END OF ACTUAL DATA ***
    # *** FAKE DATA FOR TESTING PURPOSES ***
    # yt_video_id = "ZFntEFXKDHM"
    # video_title = "Kate Darling: Social Robots, Ethics, Privacy and the Future of MIT | Lex Fridman Podcast #329"
    # episode_number = "329"
    # yt_video_date = '2022-10-15'
    # *** END OF FAKE DATA ***
    return (episode_number, yt_video_id, video_title, yt_video_date)


def prep_comment_ws(ws, wb):
    ws.title = "Positive Comments"
    ws.column_dimensions["A"].width = 15
    ws["A1"].fill = PatternFill(start_color="FFFF00", fill_type="solid")
    ws["A1"] = "Date"
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A1"].font = Font(bold=True)
    ws["A2"].fill = PatternFill(fill_type="none")
    ws.column_dimensions["B"].width = 15
    ws["B1"].fill = PatternFill(start_color="FFFF00", fill_type="solid")
    ws["B1"].value = "Source"
    ws["B1"].alignment = Alignment(horizontal="center")
    ws["B1"].font = Font(bold=True)
    ws.column_dimensions["C"].width = 15
    ws["C1"].fill = PatternFill(start_color="FFFF00", fill_type="solid")
    ws["C1"].value = "Sentiment"
    ws["C1"].alignment = Alignment(horizontal="center")
    ws["C1"].font = Font(bold=True)
    ws.column_dimensions["D"].width = 10
    ws["D1"].fill = PatternFill(start_color="FFFF00", fill_type="solid")
    ws["D1"] = "Likes"
    ws["D1"].alignment = Alignment(horizontal="center")
    ws["D1"].font = Font(bold=True)
    ws.column_dimensions["E"].width = 20
    ws["E1"].fill = PatternFill(start_color="FFFF00", fill_type="solid")
    ws["E1"] = "Username"
    ws["E1"].alignment = Alignment(horizontal="center")
    ws["E1"].font = Font(bold=True)
    ws.column_dimensions["F"].width = 200
    ws["F1"].fill = PatternFill(start_color="FFFF00", fill_type="solid")
    ws["F1"] = "Comment"
    ws["F1"].font = Font(bold=True)
    # ws.column_dimensions['G'].width = 75
    # ws['G1'].fill = PatternFill(start_color="FFFF00", fill_type="solid")
    # ws['G1'] = "Replies"
    # ws['G1'].font = Font(bold=True)
    neutral_ws = wb.copy_worksheet(ws)
    neutral_ws.title = "Neutral Comments"
    negative_ws = wb.copy_worksheet(ws)
    negative_ws.title = "Negative Comments"


def get_yt_comments(youtube, yt_video_id, ws, token, x, row_number):
    request = youtube.commentThreads().list(
        part="snippet,replies", videoId=yt_video_id, pageToken=token
    )
    search_response = request.execute()
    y = 0
    for item in search_response["items"]:
        username = item["snippet"]["topLevelComment"]["snippet"]["authorDisplayName"]
        date = item["snippet"]["topLevelComment"]["snippet"]["publishedAt"][0:10]
        comment = item["snippet"]["topLevelComment"]["snippet"]["textDisplay"]
        like_count = item["snippet"]["topLevelComment"]["snippet"]["likeCount"]
        if username != "Lex Fridman":
            sentiment_score = score_comment(comment)
        # reply_count = item['snippet']['totalReplyCount']
        # replies = []
        # x = 1
        # if reply_count > 0:
        #     print(item['replies']['comments'])
        #     for reply in item['replies']['comments']:
        #         reply_text = reply['snippet']['textDisplay']
        #         replyAuthor = reply['snippet']['authorDisplayName']
        #         replyLikes = reply['snippet']['likeCount']
        #         formattedReply = "Reply #" + str(x) + " (" + str(replyLikes) + " Likes): " + reply_text + " - " + replyAuthor
        #         replies.append(formattedReply)
        #         x = x + 1
        #     replies = '\n'.join([str(elem) for elem in replies])
        if username != "Lex Fridman":
            cell_a = "A" + str(row_number)
            cell_b = "B" + str(row_number)
            cell_c = "C" + str(row_number)
            cell_d = "D" + str(row_number)
            cell_e = "E" + str(row_number)
            cell_f = "F" + str(row_number)
            cell_g = "G" + str(row_number)
            ws[cell_a] = date
            ws[cell_a].alignment = Alignment(horizontal="center")
            ws[cell_b] = "YouTube"
            ws[cell_b].alignment = Alignment(horizontal="center")
            ws[cell_b].font = Font(color="FF0000")
            ws[cell_c] = sentiment_score
            ws[cell_c].alignment = Alignment(horizontal="center")
            ws[cell_d] = like_count
            ws[cell_d].alignment = Alignment(horizontal="center")
            ws[cell_e] = username
            ws[cell_e].alignment = Alignment(horizontal="center")
            ws[cell_f] = comment
            ws[cell_f].alignment = Alignment(wrap_text=True, shrink_to_fit=False)
            # if len(replies) != 0:
            #     ws[cell_g] = replies
            #     ws[cell_g].alignment = Alignment(wrap_text=True, shrink_to_fit=False)
            row_number = row_number + 1
        x = x + 1
    if "nextPageToken" in search_response:
        token = search_response["nextPageToken"]
        return get_yt_comments(youtube, yt_video_id, ws, token, x, row_number)


# def get_reddit_comments(video_title, ws):
#     ## Reddit Credentials
#     reddit = praw.Reddit(
#         client_id=reddit_client_id,
#         client_secret=reddit_client_secret,
#         user_agent="post bot",
#         username=reddit_username,
#         password=reddit_password,
#     )
#     subreddit_name = reddit.subreddit("lexfridman")
#     post_search = subreddit_name.search(query=video_title, limit=1)
#     for post in post_search:
#         post_id = post
#         found_post_title = post.title
#     if found_post_title != video_title:
#         print("Error: no reddit post matches the YouTube video title")
#     else:
#         row_number = ws.max_row + 1
#         reddit_comments = reddit.submission(post_id).comments
#         for comment in reddit_comments:
#             username = str(comment.author)
#             date = datetime.utcfromtimestamp(comment.created_utc).strftime(
#                 "%Y-%m-%d %H:%M:%S"
#             )[0:10]
#             text = comment.body
#             like_count = comment.ups
#             # dislike_count = comment.downs
#             sentiment_score = score_comment(text)
#             cell_a = "A" + str(row_number)
#             cell_b = "B" + str(row_number)
#             cell_c = "C" + str(row_number)
#             cell_d = "D" + str(row_number)
#             cell_e = "E" + str(row_number)
#             cell_f = "F" + str(row_number)
#             cell_g = "G" + str(row_number)
#             ws[cell_a] = date
#             ws[cell_a].alignment = Alignment(horizontal="center")
#             ws[cell_b] = "Reddit"
#             ws[cell_b].alignment = Alignment(horizontal="center")
#             ws[cell_b].font = Font(color="FF5700")
#             ws[cell_c] = sentiment_score
#             ws[cell_c].alignment = Alignment(horizontal="center")
#             ws[cell_d] = like_count
#             ws[cell_d].alignment = Alignment(horizontal="center")
#             ws[cell_e] = username
#             ws[cell_e].alignment = Alignment(horizontal="center")
#             ws[cell_f] = text
#             row_number = row_number + 1


# def get_twitter_comments(video_title, yt_video_date, ws):
#     # Authenticate account with twitter api
#     auth = tweepy.OAuthHandler(twitter_api_key, twitter_api_key_secret)
#     auth.set_access_token(twitter_access_token, twitter_access_token_secret)
#     api = tweepy.API(auth)
#     user = api.get_user(screen_name="lexfridman")
#     podcast_guest = video_title.partition(":")[0]
#     yt_video_date = datetime.strptime(yt_video_date, "%Y-%m-%d").date()
#     tweet_message = "Here's my conversation with " + podcast_guest
#     twitter_api_max_date = date.today() - timedelta(days=6)
#     if yt_video_date < twitter_api_max_date:
#         print(
#             "Error: twitter does not store date for more than a week -> cannot provide twitter comments for this episode"
#         )
#     else:
#         search_results = api.search_tweets(
#             q=tweet_message,
#             count=1,
#             result_type="popular",
#             lang="en",
#             include_entities=True,
#         )
#         tweet_id = search_results[0].id
#         replies = []
#         for tweet in tweepy.Cursor(
#             api.search_tweets, q="to:" + "lexfridman", result_type="recent"
#         ).items(10):
#             if hasattr(tweet, "in_reply_to_status_id_str"):
#                 if tweet.in_reply_to_status_id_str == tweet_id:
#                     replies.append(tweet)
#     for tweet in replies:
#         print("To Do")


def score_comment(comment):
    classifier = pipeline(
        "sentiment-analysis", model="cardiffnlp/twitter-roberta-base-sentiment"
    )
    classified = classifier(comment)
    comment_score = classified[0]["label"]
    if comment_score == "LABEL_0":
        return "Negative"
    elif comment_score == "LABEL_1":
        return "Neutral"
    elif comment_score == "LABEL_2":
        return "Positive"
    else:
        return "ERROR: could not classify comment"


def sort_comments_by_sentiment(wb, ws):
    sheets = wb.sheetnames
    neutral_ws = wb[sheets[1]]
    negative_ws = wb[sheets[2]]
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row))
    rows = reversed(rows)
    current_neutral_row = 2
    current_negative_row = 2
    total_positive = 0
    total_neutral = 0
    total_negative = 0
    for row in rows:
        sentiment_score = row[2].value
        if sentiment_score == "Neutral":
            total_neutral += 1
            cell_a = "A" + str(current_neutral_row)
            cell_b = "B" + str(current_neutral_row)
            cell_c = "C" + str(current_neutral_row)
            cell_d = "D" + str(current_neutral_row)
            cell_e = "E" + str(current_neutral_row)
            cell_f = "F" + str(current_neutral_row)
            # cell_g = 'G' + str(current_neutral_row)
            neutral_ws[cell_a] = row[0].value
            neutral_ws[cell_a].alignment = Alignment(horizontal="center")
            neutral_ws[cell_b] = row[1].value
            neutral_ws[cell_b].alignment = Alignment(horizontal="center")
            neutral_ws[cell_c] = row[2].value
            neutral_ws[cell_c].alignment = Alignment(horizontal="center")
            neutral_ws[cell_d] = row[3].value
            neutral_ws[cell_d].alignment = Alignment(horizontal="center")
            neutral_ws[cell_e] = row[4].value
            neutral_ws[cell_e].alignment = Alignment(horizontal="center")
            neutral_ws[cell_f] = row[5].value
            # neutral_ws[cell_g] = row[6].value
            if neutral_ws[cell_b].value == "YouTube":
                neutral_ws[cell_b].font = Font(color="FF0000")
            # if neutral_ws[cell_b].value == "Reddit":
            #     neutral_ws[cell_b].font = Font(color="FF5700")
            # if neutral_ws[cell_b].value == "Twitter":
            #     neutral_ws[cell_b].font = Font(color="1DA1F2")
            current_neutral_row += 1
            ws.delete_rows(row[2].row, 1)
        elif sentiment_score == "Negative":
            total_negative += 1
            cell_a = "A" + str(current_negative_row)
            cell_b = "B" + str(current_negative_row)
            cell_c = "C" + str(current_negative_row)
            cell_d = "D" + str(current_negative_row)
            cell_e = "E" + str(current_negative_row)
            cell_f = "F" + str(current_negative_row)
            # cell_g = 'G' + str(current_negative_row)
            negative_ws[cell_a] = row[0].value
            negative_ws[cell_a].alignment = Alignment(horizontal="center")
            negative_ws[cell_b] = row[1].value
            negative_ws[cell_b].alignment = Alignment(horizontal="center")
            negative_ws[cell_c] = row[2].value
            negative_ws[cell_c].alignment = Alignment(horizontal="center")
            negative_ws[cell_d] = row[3].value
            negative_ws[cell_d].alignment = Alignment(horizontal="center")
            negative_ws[cell_e] = row[4].value
            negative_ws[cell_e].alignment = Alignment(horizontal="center")
            negative_ws[cell_f] = row[5].value
            # negative_ws[cell_g] = row[6].value
            if negative_ws[cell_b].value == "YouTube":
                negative_ws[cell_b].font = Font(color="FF0000")
            # elif negative_ws[cell_b].value == "Reddit":
            #     negative_ws[cell_b].font = Font(color="FF5700")
            # elif negative_ws[cell_b].value == "Twitter":
            #     negative_ws[cell_b].font = Font(color="1DA1F2")
            current_negative_row += 1
            ws.delete_rows(row[2].row, 1)
        else:
            total_positive += 1
    return (total_positive, total_neutral, total_negative)


# Run comment analyzer 2
if __name__ == "__main__":
    main()
