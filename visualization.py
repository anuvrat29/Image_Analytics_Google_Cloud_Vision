import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import xlsxwriter,os,re,xlrd,pylab
from matplotlib import font_manager as fm
from wordcloud import WordCloud,STOPWORDS
import xlwings as xw

wb = xw.Book.caller()

def process_text(text):
    text = re.sub(r'http\S+',"", text, flags=re.MULTILINE)
    return text

def rgb_func(word, font_size, position,orientation,random_state=None, **kwargs):
    return("rgb({0},{1},{2})".format(np.random.randint(50,150),np.random.randint(0,255),np.random.randint(0,255)))

def run_visualize():
    check = "\u2714"
    wrong = "\u2716"
    filepath = os.path.join(os.path.expanduser("~"), "Downloads/Image Analytics")
    workbook=xlsxwriter.Workbook(filepath+"/Image Analytics Data Visualization.xlsx")
    worksheet=workbook.add_worksheet("Dashboard")
    worksheet.hide_gridlines(2)

    merge_format_title = workbook.add_format({'bold': 1,'font_size':20,'font':"Georgia",'border': 10,'align': 'center','valign': 'vcenter',})
    merge_format = workbook.add_format({'bold': 1,'font_size':10,'font':"Georgia",'border': 1,'align': 'center','valign': 'vcenter',})
    
    worksheet.merge_range('C2:V5', 'Dashboard for Image Analytics', merge_format_title)

    worksheet.merge_range('C7:K8', 'Image and Tweets Statistics', merge_format)
    worksheet.merge_range('C26:K27', 'Label Analysis', merge_format)
    worksheet.merge_range('C53:K54', 'Safe Search Properties', merge_format)
    worksheet.merge_range('C72:K73', 'Word Cloud of Tweets', merge_format)
    worksheet.merge_range('C90:K91', 'Logo Analysis', merge_format)

    worksheet.merge_range('N7:V8', 'Statistics of Collected Images', merge_format)
    worksheet.merge_range('N26:V27', 'Web Search Properties', merge_format)
    worksheet.merge_range('N53:V54', 'Facial Expression', merge_format)
    worksheet.merge_range('N72:V73', 'Word Cloud of Image Text', merge_format)
    worksheet.merge_range('N90:V91', 'Landmark Analysis', merge_format)

    plt.style.use("seaborn")
    plt.rcParams["font.family"] = "Georgia"

    label = pd.read_excel(filepath+"/ImageDescription.xlsx",sheet_name="Label")
    web = pd.read_excel(filepath+"/ImageDescription.xlsx",sheet_name="Web Search")
    safe = pd.read_excel(filepath+"/ImageDescription.xlsx",sheet_name="Safe Search")
    face = pd.read_excel(filepath+"/ImageDescription.xlsx",sheet_name="Face")
    logo = pd.read_excel(filepath+"/ImageDescription.xlsx",sheet_name="Logo")
    landmark = pd.read_excel(filepath+"/ImageDescription.xlsx",sheet_name="Landmark")
    text_img = pd.read_excel(filepath+"/ImageDescription.xlsx",sheet_name="Text")

    plt.figure(figsize=(11,6))
    try:
        stats = pd.read_excel(filepath+"/ImageDescription.xlsx",sheet_name="Stats")
        x = []
        if "No of Tweets" in stats.columns:
            x = ("No of Tweets","Total Images","Unique Images")
            stats = pd.Series.from_array([int(stats["No of Tweets"].values),int(stats["Total Images"].values),int(stats["Unique Images"].values)])
        else:
            x = ("No of Links","Total Images","Unique Images")
            stats = pd.Series.from_array([int(stats["No of Links"].values),int(stats["Total Images"].values),int(stats["Unique Images"].values)])
        ax = stats.plot(kind="bar",fontsize=17,rot=0)
        ax.set_xticklabels(x)
        for bar in ax.patches:
            ax.annotate(str(bar.get_height()), (bar.get_x() + bar.get_width() / 2, bar.get_height()), ha='center', va='bottom', fontsize=14)
        plt.ylabel("No of occurrences",fontsize=17)
        plt.yticks(fontsize=14)
        plt.savefig(filepath+"/stats.png",bbox_inches='tight')
        worksheet.insert_image("C10",filepath+"/stats.png",{'x_scale': 0.63, 'y_scale': 0.62})
        wb.sheets[0].range('E27').value = check
    except xlrd.biffh.XLRDError:
        plt.plot([],[])
        plt.text(-0.06, 0,"In Local System Image Analytics Real Tweet Stats not available",fontdict={'family': 'Georgia','color':'red','weight': 'normal','size': 25})
        plt.axis("off")
        plt.savefig(filepath+"/stats.png",bbox_inches='tight',facecolor="#EEEEEE")
        worksheet.insert_image("C10",filepath+"/stats.png",{'x_scale': 0.61, 'y_scale': 0.62})
        wb.sheets[0].range('E27').value = wrong
        pass

    plt.figure(figsize=(11,6))
    x = ("Label","Face","Logo","Landmark","Text","Web Search","Safe Search")
    stats_img = pd.Series.from_array([label["Label Names"].count(),face["File Name"].count(),logo["Logo Names"].count(),landmark["Landmark Names"].count(),text_img["Text"].count(),web["Web Search Properties"].count(),safe["File Name"].count()])
    ax = stats_img.plot(kind="bar",fontsize=14,rot=0)
    ax.set_xticklabels(x)
    for bar in ax.patches:
        ax.annotate(str(bar.get_height()), (bar.get_x() + bar.get_width() / 2, bar.get_height()), ha='center', va='bottom', fontsize=14)
    plt.ylabel("No of occurrences",fontsize=17)
    plt.yticks(fontsize=14)
    plt.savefig(filepath+"/stats_properties.png",bbox_inches='tight')
    worksheet.insert_image("N10",filepath+"/stats_properties.png",{'x_scale': 0.65, 'y_scale': 0.62})

    plt.figure(figsize=(11,6))
    ax = label["Label Names"].value_counts().nlargest(10).plot(kind="bar")
    plt.xticks(fontsize=16,rotation=65)
    for bar in ax.patches:
        ax.annotate(str(bar.get_height()), (bar.get_x() + bar.get_width() / 2, bar.get_height()), ha='center', va='bottom', fontsize=14)
    plt.yticks(fontsize=14)
    plt.ylabel("No of occurrences",fontsize=18)
    plt.savefig(filepath+"/label.png",bbox_inches='tight')
    worksheet.insert_image("C29",filepath+"/label.png",{'x_scale': 0.64, 'y_scale': 0.62})
    wb.sheets[0].range('F27').value = check

    plt.figure(figsize=(11,6))
    ax = web["Web Search Properties"].value_counts().nlargest(10).plot(kind="bar")
    plt.xticks(fontsize=16,rotation=65)
    for bar in ax.patches:
        ax.annotate(str(bar.get_height()), (bar.get_x() + bar.get_width() / 2, bar.get_height()), ha='center', va='bottom', fontsize=14)
    plt.yticks(fontsize=14)
    plt.ylabel("No of occurrences",fontsize=18)
    plt.savefig(filepath+"/websearch.png",bbox_inches='tight')
    worksheet.insert_image("N29",filepath+"/websearch.png",{'x_scale': 0.65, 'y_scale': 0.62})
    wb.sheets[0].range('G27').value = check

    sw= set(STOPWORDS)
    plt.figure(figsize=(13,6))
    try:
        content = pd.read_excel(filepath+"/ImageDescription.xlsx",sheet_name="Text Analytics")
        text = process_text("".join(content["CONTENT"]))
        tweet = WordCloud(font_path=fm.findfont("Georgia"),background_color="#EEEEEE",max_words=2000,normalize_plurals= True,stopwords=sw,
                                                  width=1500, height=750).generate(text=text)
        tweet.recolor(color_func=rgb_func)
        plt.imshow(tweet)
        plt.axis("off")
        plt.savefig(filepath+"/tweet.png",bbox_inches='tight',facecolor="#EEEEEE")
        worksheet.insert_image("C75",filepath+"/tweet.png",{'x_scale': 0.61, 'y_scale': 0.58})
        wb.sheets[0].range('H27').value = check
    except xlrd.biffh.XLRDError:
        plt.plot([],[])
        plt.text(-0.05, 0,"In Local System Image Analytics Tweets not available",fontdict={'family': 'Georgia','color':'red','weight': 'normal','size': 25})
        plt.axis("off")
        plt.savefig(filepath+"/tweet.png",bbox_inches='tight',facecolor="#EEEEEE")
        worksheet.insert_image("C75",filepath+"/tweet.png",{'x_scale': 0.56, 'y_scale': 0.59})
        wb.sheets[0].range('H27').value = wrong
        pass

    plt.figure(figsize=(13,6))
    if not text_img.empty:
        text = process_text("".join(text_img["Text"]))
        img_text = WordCloud(font_path=fm.findfont("Georgia"),background_color="#EEEEEE",max_words=2000,normalize_plurals= True,stopwords=sw,
                                                         width=1500, height=750).generate(text=text)
        img_text.recolor(color_func=rgb_func)
        plt.imshow(img_text)
        plt.axis("off")
        plt.savefig(filepath+"/img_text.png",bbox_inches='tight',facecolor="#EEEEEE")
        worksheet.insert_image("N75",filepath+"/img_text.png",{'x_scale': 0.61, 'y_scale': 0.58})
        wb.sheets[0].range('I27').value = check
    else:
        plt.pie([],labels=[])
        plt.text(-0.35, 0,"Data Not Available",fontdict={'family': 'Georgia','color':'red','weight': 'normal','size': 25})
        plt.axis("off")
        plt.savefig(filepath+"/img_text.png",bbox_inches='tight',facecolor="#EEEEEE")
        worksheet.insert_image("N75",filepath+"/img_text.png",{'x_scale': 0.58, 'y_scale': 0.59})
        wb.sheets[0].range('I27').value = wrong

    plt.figure(figsize=(11,6))
    plt.subplot(2,2,1)
    safe["Adult"].value_counts().plot(kind="pie",autopct='%1.1f%%',startangle=0,fontsize=14)
    pylab.ylabel('')
    plt.title("Adult",fontsize=15,fontweight="bold")
    plt.axis("equal")
    plt.subplot(2,2,2)
    safe["Medical"].value_counts().plot(kind="pie",autopct='%1.1f%%',startangle=0,fontsize=14)
    pylab.ylabel('')
    plt.title("Medical",fontsize=15,fontweight="bold")
    plt.axis("equal")
    plt.subplot(2,2,3)
    safe["Spoof"].value_counts().plot(kind="pie",autopct='%1.1f%%',startangle=0,fontsize=14)
    pylab.ylabel('')
    plt.title("Spoof",fontsize=15,fontweight="bold")
    plt.axis("equal")
    plt.subplot(2,2,4)
    safe["Violence"].value_counts().plot(kind="pie",autopct='%1.1f%%',startangle=0,fontsize=14)
    plt.title("Violence",fontsize=15,fontweight="bold")
    pylab.ylabel('')
    plt.axis("equal")
    plt.savefig(filepath+"/safe.png",bbox_inches='tight',facecolor="#EEEEEE")
    worksheet.insert_image("C56",filepath+"/safe.png",{'x_scale': 0.69, 'y_scale': 0.62})
    wb.sheets[0].range('J27').value = check

    plt.figure(figsize=(11,6))
    if not face.empty:
        plt.subplot(221)
        face["Anger"].value_counts().plot(kind="pie",autopct='%1.1f%%',startangle=0,fontsize=14)
        pylab.ylabel('')
        plt.title("Anger",fontsize=15,fontweight="bold")
        plt.axis("equal")
        plt.subplot(222)
        face["Joy"].value_counts().plot(kind="pie",autopct='%1.1f%%',startangle=0,fontsize=14)
        pylab.ylabel('')
        plt.title("Joy",fontsize=15,fontweight="bold")
        plt.axis("equal")
        plt.subplot(212)
        face["Surprise"].value_counts().plot(kind="pie",autopct='%1.1f%%',startangle=0,fontsize=14)
        pylab.ylabel('')
        plt.title("Surprise",fontsize=15,fontweight="bold")
        plt.axis("equal")
        plt.savefig(filepath+"/face.png",bbox_inches='tight',facecolor="#EEEEEE")
        worksheet.insert_image("N56",filepath+"/face.png",{'x_scale': 0.69, 'y_scale': 0.61})
        wb.sheets[0].range('K27').value = check
    else:
        plt.pie([],labels=[])
        plt.text(-0.35, 0,"Data Not Available",fontdict={'family': 'Georgia','color':'red','weight': 'normal','size': 25})
        plt.savefig(filepath+"/face.png",bbox_inches='tight',facecolor="#EEEEEE")
        worksheet.insert_image("N56",filepath+"/face.png",{'x_scale': 0.69, 'y_scale': 0.62})
        wb.sheets[0].range('K27').value = wrong
    
    plt.figure(figsize=(11,6))
    if not logo.empty:
        ax = logo["Logo Names"].value_counts().nlargest(10).plot(kind="bar")
        plt.xticks(fontsize=16,rotation=65)
        for bar in ax.patches:
            ax.annotate(str(bar.get_height()), (bar.get_x() + bar.get_width() / 2, bar.get_height()), ha='center', va='bottom', fontsize=14)
        plt.yticks(fontsize=14)
        plt.ylabel("No of occurrences",fontsize=18)
        plt.savefig(filepath+"/logo.png",bbox_inches='tight',facecolor="#EEEEEE")
        worksheet.insert_image("C93",filepath+"/logo.png",{'x_scale': 0.64, 'y_scale': 0.63})
        wb.sheets[0].range('L27').value = check
    else:
        plt.plot([],[])
        plt.axis("off")
        plt.text(-0.022, 0,"Data Not Available",fontdict={'family': 'Georgia','color':'red','weight': 'normal','size': 25})
        plt.savefig(filepath+"/logo.png",bbox_inches='tight',facecolor="#EEEEEE")
        worksheet.insert_image("C93",filepath+"/logo.png",{'x_scale': 0.65, 'y_scale': 0.64})
        wb.sheets[0].range('L27').value = wrong

    plt.figure(figsize=(11,6))
    if not landmark.empty:
        ax = landmark["Landmark Names"].value_counts().nlargest(10).plot(kind="bar")
        plt.xticks(fontsize=16,rotation=65)
        for bar in ax.patches:
            ax.annotate(str(bar.get_height()), (bar.get_x() + bar.get_width() / 2, bar.get_height()), ha='center', va='bottom', fontsize=14)
        plt.yticks(fontsize=14)
        plt.ylabel("No of occurrences",fontsize=18)
        plt.savefig(filepath+"/landmark.png",bbox_inches='tight',facecolor="#EEEEEE")
        worksheet.insert_image("N93",filepath+"/landmark.png",{'x_scale': 0.64, 'y_scale': 0.63})
        wb.sheets[0].range('M27').value = check
    else:
        plt.plot([],[])
        plt.axis("off")
        plt.text(-0.022, 0,"Data Not Available",fontdict={'family': 'Georgia','color':'red','weight': 'normal','size': 25})
        plt.savefig(filepath+"/landmark.png",bbox_inches='tight',facecolor="#EEEEEE")
        worksheet.insert_image("N93",filepath+"/landmark.png",{'x_scale': 0.65, 'y_scale': 0.64})
        wb.sheets[0].range('M27').value = wrong

    workbook.close()

    try:
        os.remove(filepath+"/stats.png")
        os.remove(filepath+"/stats_properties.png")
        os.remove(filepath+"/label.png")
        os.remove(filepath+"/websearch.png")
        os.remove(filepath+"/tweet.png")
        os.remove(filepath+"/img_text.png")
        os.remove(filepath+"/safe.png")
        os.remove(filepath+"/face.png")
        os.remove(filepath+"/logo.png")
        os.remove(filepath+"/landmark.png")
    except OSError:
        pass