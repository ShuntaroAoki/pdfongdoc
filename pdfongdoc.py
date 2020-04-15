#!/usr/bin/env python
# coding: utf-8


import argparse
import os
import pickle
from pprint import pprint

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from pdf2image import convert_from_path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('document_id', metavar='DOCID', type=str, nargs=1, help='Google Doc document ID.')
    parser.add_argument('pdfs', metavar='PDF', type=str, nargs='+', help='PDF file(s) to be pasted.')
    args = parser.parse_args()

    target_doc_id = args.document_id[0]
    pdf_files = args.pdfs

    figures = pdf2jpg(pdf_files)

    # Google API auth
    creds = google_auth()
    drive = build('drive', 'v3', credentials=creds)
    gdoc = build('docs', 'v1', credentials=creds)

    # Get target doc info
    target_doc_info = drive.files().get(fileId=target_doc_id, fields='name,parents').execute()
    target_doc_name = target_doc_info['name']
    target_dir_id = target_doc_info['parents'][0]

    print('Document: {}'.format(target_doc_name))

    # Create upload dir
    upload_dir_id = create_dir(drive, target_doc_name + '.resources', target_dir_id)

    for figname, fig in figures.items():
        print('Figure: {}'.format(figname))

        # Upload PDF
        print('Uploading {}'.format(fig['pdf']))
        pdf_id = upload_file(drive, fig['pdf'], os.path.basename(fig['pdf']), upload_dir_id, 'application/pdf')
        pdf_url = 'https://drive.google.com/open?id=' + pdf_id

        for img in fig['image']:
            # Upload image
            print('Uploading {}'.format(img))
            img_id = upload_file(drive, img, os.path.basename(img), upload_dir_id, 'image/jpeg', reader='anyone')
            image_uri = 'https://drive.google.com/uc?export=view&id=' + img_id

            # Paste image on gdoc
            print('Pasting image')
            paste_image_pdf(gdoc, target_doc_id, image_uri, pdf_url)

    return None


def pdf2jpg(pdf_files):
    figures = {}
    for pdf_file in pdf_files:
        image_name = os.path.splitext(os.path.basename(pdf_file))[0]
        images = convert_from_path(pdf_file)
        image_files = []
        for i, image in enumerate(images):
            fpath = os.path.join(os.path.dirname(pdf_file), image_name + '_%d' % i + '.jpg')
            image.save(fpath)
            image_files.append(fpath)
        figures.update({image_name: {'pdf': pdf_file, 'image': image_files}})
    return figures


def google_auth():
    confdir = os.path.join(os.path.expanduser("~"), '.pdfongdoc')

    creds_pickle = os.path.join(confdir, 'credentials.pkl')
    creds = None
    scopes = ['https://www.googleapis.com/auth/drive',
              'https://www.googleapis.com/auth/documents']

    if os.path.exists(creds_pickle):
        with open(creds_pickle, 'rb') as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(os.path.join(confdir, 'client_secrets.json'), scopes)
            creds = flow.run_local_server(port=0)
            with open(creds_pickle, 'wb') as f:
                pickle.dump(creds, f, protocol=2)

    return creds


def create_dir(drive, name, parent_id):
    resp = drive.files().create(
        body={
            'name': name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_id]
        },
        fields='id'
    ).execute()

    return resp['id']


def upload_file(drive, fpath, name, parent_id, mimetype, reader=None):
    media_body = MediaFileUpload(fpath, mimetype=mimetype)
    res = drive.files().create(
        body={
            'name': name,
            'mimetype': mimetype,
            'parents': [parent_id]
        },
        media_body=media_body,
        fields='id'
    ).execute()

    if reader == 'anyone':
        drive.permissions().create(
            fileId=res['id'],
            body={
                'role': 'reader',
                'type': 'anyone'
            }
        ).execute()

    return res['id']


def paste_image_pdf(gdoc, doc_id, image_uri, pdf_url):

    doc = gdoc.documents().get(
        documentId=doc_id
    ).execute()
    imagepos = max([c['endIndex'] for c in doc['body']['content']]) - 1

    requests = [
        {
            'insertInlineImage': {
                'location': {
                    'index': imagepos
                },
                'uri': image_uri
            }
        },
        {
            'updateTextStyle': {
                'textStyle': {
                    'link' : {
                        'url': pdf_url
                    }
                },
                'range': {
                    'startIndex': imagepos,
                    'endIndex': imagepos + 1
                },
                'fields': 'link'
            }
        }
    ]
    body = {'requests': requests}
    resp = gdoc.documents().batchUpdate(
        documentId=doc_id, body=body
    ).execute()
    return None


if __name__ == '__main__':
    main()
