import pandas as pd
import json
import os

# Define file paths
data_dir = os.path.join(os.path.dirname(__file__), 'data')
content_file = os.path.join(data_dir, 'african-sap-user-group_content_1768374674960.xls')
followers_file = os.path.join(data_dir, 'african-sap-user-group_followers_1768374852794.xls')
competitor_file = os.path.join(data_dir, 'African SAP User Group (AFSUG)_competitor_analytics_1768374945515.xlsx')
visitors_file = os.path.join(data_dir, 'african-sap-user-group_visitors_1768374824893.xls')

# 1. KPIs and Trends from Competitor and Metrics
comp_df = pd.read_excel(competitor_file, engine='openpyxl', header=1)
afsug_comp = comp_df[comp_df['Page'].str.contains('AFSUG', na=False)].iloc[0]

total_followers = int(afsug_comp['Total Followers'])
new_followers_period = int(afsug_comp['New Followers'])

# Daily metrics
content_metrics = pd.read_excel(content_file, sheet_name='Metrics', header=1)
# Ensure Date is string for JSON
content_metrics['Date'] = content_metrics['Date'].astype(str)

# Followers daily
followers_daily = pd.read_excel(followers_file, sheet_name=0, header=0)
followers_daily['Date'] = followers_daily['Date'].astype(str)

# 2. Top Posts
top_posts_df = pd.read_excel(content_file, sheet_name='All posts', header=1)
# Sort by Engagement rate and take top 5
top_5_posts = top_posts_df.sort_values(by='Engagement rate', ascending=False).head(5)
top_5_posts_list = top_5_posts[['Post title', 'Engagement rate']].to_dict(orient='records')

# 3. Demographics
xl_followers = pd.ExcelFile(followers_file)
demographics = {}
for sheet in ['Location', 'Job function', 'Seniority', 'Industry']:
    df = pd.read_excel(xl_followers, sheet_name=sheet)
    # Get top 5 or so for each
    top_dem = df.head(10).to_dict(orient='records')
    demographics[sheet] = top_dem

# Calculate summary stats
summary = {
    "total_followers": total_followers,
    "follower_growth_pct": round((new_followers_period / (total_followers - new_followers_period)) * 100, 1),
    "total_impressions": int(content_metrics['Impressions (total)'].sum()),
    "total_engagements": int(content_metrics['Reactions (total)'].sum() + content_metrics['Comments (total)'].sum() + content_metrics['Reposts (total)'].sum() + content_metrics['Clicks (total)'].sum()),
}
summary["engagement_rate"] = round((summary["total_engagements"] / summary["total_impressions"]) * 100, 2) if summary["total_impressions"] > 0 else 0

# Prepare final data
data = {
    "summary": summary,
    "trends": {
        "dates": content_metrics['Date'].tolist(),
        "impressions": content_metrics['Impressions (total)'].tolist(),
        "engagement": (content_metrics['Reactions (total)'] + content_metrics['Comments (total)'] + content_metrics['Reposts (total)'] + content_metrics['Clicks (total)']).tolist(),
        "followers_daily": followers_daily['Total followers'].tolist(), # This is daily new followers
        "follower_dates": followers_daily['Date'].tolist()
    },
    "top_posts": top_5_posts_list,
    "demographics": demographics
}

# Add cumulative followers for trend
current_f = total_followers
cumulative_followers = []
# Work backwards from total_followers
new_f_daily = followers_daily['Total followers'].tolist()
for val in reversed(new_f_daily):
    cumulative_followers.append(current_f)
    current_f -= val
data["trends"]["followers_cumulative"] = list(reversed(cumulative_followers))

output_file = os.path.join(os.path.dirname(__file__), 'data.json')
with open(output_file, 'w') as f:
    json.dump(data, f, indent=4)

print(f"Data processed successfully and saved to {output_file}")

# Also embed data into index.html to avoid CORS issues for local viewing
index_file = os.path.join(os.path.dirname(__file__), 'index.html')
if os.path.exists(index_file):
    with open(index_file, 'r') as f:
        html_content = f.read()
    
    data_json = json.dumps(data, indent=4)
    data_script = f"const dashboardData = {data_json};"
    
    new_html = html_content.replace('// DATA_PLACEHOLDER', data_script)
    
    with open(index_file, 'w') as f:
        f.write(new_html)
    print(f"Data successfully embedded into {index_file}")
