import * as React from 'react';
import { DefaultButton, Dialog, DialogType, DialogFooter, TextField, Dropdown, IDropdownOption, PrimaryButton, ProgressIndicator } from '@fluentui/react';
import { SPHttpClient,  SPHttpClientResponse} from '@microsoft/sp-http';
import { sp } from '@pnp/sp/presets/all';

interface MyDialogPopupProps {
  absoluteURL: string;
  spHttpClient: SPHttpClient;
}

const classificationOptions: IDropdownOption[] = [
  { key: 'private', text: 'Private' },
  { key: 'public', text: 'Public' },
  { key: 'confidential', text: 'Confidential' }
];

const MyDialogPopup: React.FC<MyDialogPopupProps> = ({ absoluteURL,spHttpClient}) => {
  const [isDialogOpen, setIsDialogOpen] = React.useState(false);
  const [file, setFile] = React.useState<File | null>(null);
  const [classification, setClassification] = React.useState<string | undefined>(undefined);
  const [validity, setValidity] = React.useState<number | undefined>(undefined);
  const [errorMessage, setErrorMessage] = React.useState<string | null>(null);
  const [uploadProgress, setUploadProgress] = React.useState<number>(0);
  const [isUploading, setIsUploading] = React.useState<boolean>(false);
  const [requestDigest, setRequestDigest] = React.useState<string | null>(null);

 const allowedFileTypes = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'application/x-zip-compressed'
];

  const fetchRequestDigest = async () => {
    const response: SPHttpClientResponse = await spHttpClient.post(`${absoluteURL}/_api/contextinfo`, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      }
    });
    const data = await response.json();

    setRequestDigest(data.FormDigestValue);
  };
  React.useEffect(() => {
    fetchRequestDigest();
  }, []);

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) :any => {
    const input = event.target as HTMLInputElement;
    if (input.files) {
      const selectedFile = input.files[0];
      if (allowedFileTypes.includes(selectedFile.type)) {
        setFile(selectedFile);
        setErrorMessage(null);
      } else {
        setFile(null);
        setErrorMessage('Invalid file type. Please upload an .xlsx, .docx, or .pptx file.');
      }
    }
  };

  const handleClassificationChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    setClassification(option?.key as string);
  };

  const handleValidityChange = (event: React.FormEvent<HTMLInputElement>, newValue?: string) => {
    setValidity(newValue ? parseInt(newValue) : undefined);
  };

  const uploadFile=async (file: File) => {

    
  };

  const startUpload = async (file: File) => {
    const response: SPHttpClientResponse = await spHttpClient.post(`${absoluteURL}/_api/web/getfolderbyserverrelativeurl('Shared Documents')/files/add(url='${file.name}',overwrite=true)`, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json',
        'X-RequestDigest': requestDigest || ''
      }
    });
    return response.json();
  };

 const continueUpload = async (file: File, uploadId: string, offset: number, chunk: Blob,ServerRelativeUrl:string) => {
    const response: SPHttpClientResponse = await spHttpClient.post(`${absoluteURL}/_api/web/GetFileById('${uploadId}')/ContinueUpload(uploadId=guid'${uploadId}',fileOffset=${offset})`, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'X-RequestDigest': requestDigest || ''
      },
      body: chunk
    });

    await fetchRequestDigest();
    return response.json();

  };

  const finishUpload = async (file: File, uploadId: string, offset: number,ServerRelativeUrl:string) => {
    const response: SPHttpClientResponse = await spHttpClient.post(`${absoluteURL}/_api/web/getfilebyserverrelativeurl('${ServerRelativeUrl}')/finishupload(uploadId=guid'${uploadId}',fileOffset=${offset})`, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'X-RequestDigest': requestDigest || ''
      }
    });
    return response.json();
  };

  const handleSubmit = async () => {
    if (file && classification && validity) {
      try {
        setIsUploading(true);
        await fetchRequestDigest();
        const startUploadData = await startUpload(file);
        const uploadId = startUploadData.UniqueId;
        const chunkSize = 18560; // 10 MB
        let offset = 0;

        while (offset < file.size) {
          const chunk = file.slice(offset, offset + chunkSize);
          await continueUpload(file, uploadId, offset, chunk,startUploadData.ServerRelativeUrl);
          offset += chunkSize;
          setUploadProgress((offset / file.size) * 100);
        }

        await finishUpload(file, uploadId, file.size,startUploadData.ServerRelativeUrl);

        console.log('File:', file);
        console.log('Classification:', classification);
        console.log('Validity (days):', validity);
        setIsDialogOpen(false);
        setIsUploading(false);
        setUploadProgress(0);
      } catch (error) {
        setErrorMessage('Upload failed. Please try again.');
        console.error(error);
        setIsUploading(false);
      }
    } else {
      setErrorMessage('Please fill out all fields and upload a valid file.');
    }
  };

  return (
    <div>
      <DefaultButton text="Open Dialog" onClick={() => setIsDialogOpen(true)} />
      <Dialog
        hidden={!isDialogOpen}
        onDismiss={() => setIsDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: 'Upload File',
          closeButtonAriaLabel: 'Close'
        }}
        modalProps={{
          isBlocking: false
        }}
      >
        <div>
          <input type="file" accept=".xlsx,.docx,.pptx,.zip" onChange={handleFileChange} />
          {errorMessage && <p style={{ color: 'red' }}>{errorMessage}</p>}
          <Dropdown
            label="Classification"
            options={classificationOptions}
            onChange={handleClassificationChange}
          />
          <TextField
            label="Validity (days)"
            type="number"
            onChange={handleValidityChange}
          />
          {isUploading && <ProgressIndicator label="Uploading file..." percentComplete={uploadProgress / 100} />}
        </div>
        <DialogFooter>
          <PrimaryButton onClick={handleSubmit} text="Submit" />
          <DefaultButton onClick={() => setIsDialogOpen(false)} text="Cancel" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default MyDialogPopup;
