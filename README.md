# Label Generator for QBO

![alt text](https://github.com/ygalvao/label-generator-for-qbo/screenshot.png?raw=true)

## Description

This Flask-based web application is designed for generating 4x6 shipping labels using QuickBooks Online (QBO) data and it is optimized for deployment on Google Cloud Run.

## Quickstart

1. Navigate to your local repository directory.

```bash
cd your-repository-path/
```

2. Replace or edit all ".example" files.

#### To deploy this application to Google Cloud Run:

- Run the following command:

```bash
gcloud run deploy [DESIRED NAME FOR THE APP] --source . --region [DESIRED REGION] --allow-unauthenticated --memory 1G
```

Example:

```bash
gcloud run deploy label-generator --source . --region us-central1 --allow-unauthenticated --memory 1G
```

#### To run this application on your local machine:

- Run the following commands:

```bash
pip install -r requirements.txt
python app.py --on-premises [--sandbox or --production]
```

## Application Flags

### --sandbox

The --sandbox flag is used to run the application on Intuit's (QBO) sandbox environment. This is typically used for testing and development purposes, where changes can be made and tested without affecting the production environment.

### --on-premises

The --on-premises flag allows the application to be run on a local machine.

## Note on .example Files

All ".example" files provided in this repository are templates. They should be either replaced or renamed without the ".example" extension - if you choose the second option, then moddify the content with the actual values relevant to your deployment.
