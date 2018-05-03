import openpyxl,os
import xlwings as xw
import string
import pandas as pd
from google.cloud import vision
from google.cloud.vision import types

filepath = os.path.dirname(os.path.abspath(__file__))
filepath = filepath+"/apikey.json"
os.environ['GOOGLE_APPLICATION_CREDENTIALS']=filepath

wb = xw.Book.caller()
wb.sheets[0].range('L20').value = "I am processing Image Analysis Request..."

vision_client = vision.ImageAnnotatorClient()
image = types.Image()

def text_process(message):
    no_punctuation = [char for char in message if char not in string.punctuation]
    no_punctuation = ' '.join(no_punctuation)
    return no_punctuation

def run_ws_analysis():
    file_label,file_face,file_logo,file_land = [],[],[],[]
    file_text,file_web,file_safe,file_color = [],[],[],[]
    labelnames,label_scores,logonames = [],[],[]
    landmarknames,text_info,webentities = [],[],[]
    angerArr,joyArr,surpriseArr = [],[],[]
    adult,medical,spoof,violence = [],[],[],[]
    f,r,g,b = [],[],[],[]
    sample_array = []
    likelihood_name = ("Not Known","Very Less","Less","May Be","Strong","Very Strong")

    filepath = os.path.dirname(os.path.abspath(__file__))
    dataset = pd.read_excel(filepath+"/Url Data.xlsx",sheet_name=0)
    dataset_content = dataset
    dataset_stats = pd.read_excel(filepath+"/TwitterLinkUrl.xlsx",sheet_name="Stats")
    dataset = pd.read_excel(filepath+"/TwitterLinkUrl.xlsx",sheet_name="LinkUrls")
    image_urls = dataset.iloc[:,0].values

    i=0
    for uri in image_urls:
        file = uri
        image.source.image_uri = uri

        labels = vision_client.label_detection(image=image).label_annotations
        logos = vision_client.logo_detection(image=image).logo_annotations
        landmarks = vision_client.landmark_detection(image=image).landmark_annotations
        notes = vision_client.web_detection(image=image).web_detection
        texts = vision_client.text_detection(image=image).text_annotations
        faces = vision_client.face_detection(image=image).face_annotations
        safe = vision_client.safe_search_detection(image=image).safe_search_annotation
        props = vision_client.image_properties(image=image).image_properties_annotation

        for label in labels:
            file_label.append(file)
            labelnames.append(label.description)
            label_scores.append(label.score*100)

        for logo in logos:
            file_logo.append(file)
            logonames.append(logo.description)

        for landmark in landmarks:
            file_land.append(file)
            landmarknames.append(landmark.description)

        if notes.web_entities:
            for entity in notes.web_entities:
                file_web.append(file)
                webentities.append(entity.description)

        for text in texts:
            sample_array.append(text.description)
        if len(sample_array)>0:
            file_text.append(file)
            text_info.append(text_process(sample_array[1:]))
        sample_array = []

        for face in faces:
            if face:
                file_face.append(file)
                angerArr.append(likelihood_name[face.anger_likelihood])
                joyArr.append(likelihood_name[face.joy_likelihood])
                surpriseArr.append(likelihood_name[face.surprise_likelihood])

        file_safe.append(file)
        adult.append(likelihood_name[safe.adult])
        medical.append(likelihood_name[safe.medical])
        spoof.append(likelihood_name[safe.spoof])
        violence.append(likelihood_name[safe.violence])

        for color in props.dominant_colors.colors:
            file_color.append(file)
            f.append(color.pixel_fraction)
            r.append(color.color.red)
            g.append(color.color.green)
            b.append(color.color.blue)
        i=i+1
        wb.sheets[0].range('L20').value = str(i)+" image/s analyzed, "+str(image_urls.size-i)+" image/s remaining."

    wb.sheets[0].range('L20').value = "I am processed your images..."
    filepath = os.path.join(os.path.expanduser("~"), "Downloads/Image Analytics")
    if not os.path.exists(filepath):
        os.makedirs(filepath)
    writer = pd.ExcelWriter(filepath+"/ImageDescription.xlsx")

    dataset_label = pd.DataFrame({"File Name":file_label,"Label Names":labelnames,"Label Score":label_scores})
    dataset_logo = pd.DataFrame({"File Name":file_logo,"Logo Names":logonames})
    dataset_land = pd.DataFrame({"File Name":file_land,"Landmark Names":landmarknames})
    dataset_web = pd.DataFrame({"File Name":file_web,"Web Search Properties":webentities})
    dataset_text = pd.DataFrame({"File Name":file_text,"Text":text_info})
    dataset_face = pd.DataFrame({"File Name":file_face,"Anger":angerArr,"Joy":joyArr,"Surprise":surpriseArr},columns=["File Name","Anger","Joy","Surprise"])
    dataset_safe = pd.DataFrame({"File Name":file_safe,"Adult":adult,"Medical":medical,"Spoof":spoof,"Violence":violence},columns=["File Name","Adult","Medical","Spoof","Violence"])
    dataset_color = pd.DataFrame({"File Name":file_color,"Pixel Fraction":f,"Red":r,"Green":g,"Blue":b},columns=["File Name","Pixel Fraction","Red","Green","Blue"])

    dataset_label.to_excel(writer,sheet_name="Label",index=False)
    dataset_face.to_excel(writer,sheet_name="Face",index=False)
    dataset_logo.to_excel(writer,sheet_name="Logo",index=False)
    dataset_land.to_excel(writer,sheet_name="Landmark",index=False)
    dataset_text.to_excel(writer,sheet_name="Text",index=False)
    dataset_web.to_excel(writer,sheet_name="Web Search",index=False)
    dataset_safe.to_excel(writer,sheet_name="Safe Search",index=False)
    dataset_color.to_excel(writer,sheet_name="Color Gradient",index=False)
    dataset_content.to_excel(writer,sheet_name="Text Analytics",index=False)
    dataset_stats.to_excel(writer,sheet_name="Stats",index=False)

    writer.save()
    wb.sheets[0].range('L20').value = "Downloaded Analyzed file Successfully in Download/Image Analytics folder"