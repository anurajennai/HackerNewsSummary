import requests
import datetime
from newspaper import Article
import nltk
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment


# Ensure the 'punkt' tokenizer is downloaded
nltk.download('punkt')

def get_hn_stories():
    response = requests.get('https://hacker-news.firebaseio.com/v0/topstories.json?print=pretty')
    story_ids = response.json()
    return story_ids

def get_story_details(story_id):
    url = f'https://hacker-news.firebaseio.com/v0/item/{story_id}.json?print=pretty'
    response = requests.get(url)
    return response.json()

def summarize_article(url):
    try:
        article = Article(url)
        article.download()
        article.parse()
        article.nlp()
        return article.summary
    except Exception as e:
        print(f"Failed to process the URL: {url}, Error: {str(e)}")
        return None

def summarize_stories():
    story_ids = get_hn_stories()
    summaries = []
    one_day_ago = datetime.datetime.now() - datetime.timedelta(days=1)

    for story_id in story_ids[:60]:  # Limit to the first 60 stories to manage performance
        story = get_story_details(story_id)
        if story and 'time' in story:
            story_time = datetime.datetime.fromtimestamp(story['time'])
            if story_time > one_day_ago and story.get('descendants', 0) > 50:
                article_summary = summarize_article(story.get('url', ''))
                if article_summary:
                    summary = {
                        'title': story.get('title'),
                        'url': story.get('url'),
                        'comments': story.get('descendants'),
                        'summary': article_summary
                    }
                    summaries.append(summary)

    return summaries

def save_summaries_to_excel(summaries):
    df = pd.DataFrame(summaries)
    filename = datetime.datetime.now().strftime("hnsummary_%Y-%m-%d_%H%M%S.xlsx")
    # Save the DataFrame to an Excel file
    df.to_excel(filename, index=False)

    # Load the workbook and select the active worksheet
    workbook = openpyxl.load_workbook(filename)
    worksheet = workbook.active

    # Apply text wrapping
    for col in ['A', 'B', 'D']:  # Assuming A, B, D are title, URL, summary
        for cell in worksheet[col]:
            cell.alignment = Alignment(wrapText=True)

    # Define default column widths
    worksheet.column_dimensions['A'].width = 30  # Title
    worksheet.column_dimensions['B'].width = 30  # URL
    worksheet.column_dimensions['C'].width = 10  # Comments
    worksheet.column_dimensions['D'].width = 50  # Summary

    # Save the changes to the workbook
    workbook.save(filename)
    print(f"Saved summaries with formatted cells to {filename}")

if __name__ == '__main__':
    summarized_stories = summarize_stories()
    save_summaries_to_excel(summarized_stories)