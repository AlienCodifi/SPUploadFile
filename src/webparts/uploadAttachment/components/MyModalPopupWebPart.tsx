import * as React from 'react';
import { DefaultButton, Dialog, DialogType, DialogFooter, TextField, Dropdown, IDropdownOption, PrimaryButton, ProgressIndicator } from '@fluentui/react';
//import "@pnp/sp/webs";
import "@pnp/sp/context-info";
//import { IFile, IResponseItem } from "./interfaces";
import { getSP } from "../pnpconfig";
import { spfi } from "@pnp/sp";
import { Caching } from "@pnp/queryable";
import "@pnp/sp/files";
import "@pnp/sp/folders";


interface MyDialogPopupProps {
  absoluteURL: string}

const classificationOptions: IDropdownOption[] = [
  { key: 'private', text: 'Private' },
  { key: 'public', text: 'Public' },
  { key: 'confidential', text: 'Confidential' }
];


const MyDialogPopup: React.FC<MyDialogPopupProps> = ({ absoluteURL}) => {
  const [isDialogOpen, setIsDialogOpen] = React.useState(false);
  const [file, setFile] = React.useState<File | null>(null);
  const [classification, setClassification] = React.useState<string | undefined>(undefined);
  const [validity, setValidity] = React.useState<number | undefined>(undefined);
  const [errorMessage, setErrorMessage] = React.useState<string | null>(null);
  const [uploadProgress, setUploadProgress] = React.useState<number>(0);
  const [isUploading, setIsUploading] = React.useState<boolean>(false);
   const allowedFileTypes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'application/x-zip-compressed'
  ];
  const sp= spfi(getSP()).using(Caching({store:"session"}));


  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0] || null;
    if (selectedFile && !allowedFileTypes.includes(selectedFile.type)) {
      setErrorMessage('Selected file type is not allowed');
      setFile(null);
    } else {
      setErrorMessage(null);
      setFile(selectedFile);
    }
  };

  const handleClassificationChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    setClassification(option?.key as string);
  };

  const handleValidityChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    setValidity(Number(event.target.value));
  };

  const handleUpload = async () => {
    setIsUploading(true);
    if (!file) {
      setErrorMessage('No file selected');
      return;
    }
    const fileNamePath = encodeURI(file.name);
    if (!file || !classification || !validity) {
      setErrorMessage('Please fill all fields');
      return;
    }
    let result: any;
// you can adjust this number to control what size files are uploaded in chunks
if (file.size <= 10485760) {
    // small upload
    result = await sp.web.getFolderByServerRelativePath("Shared Documents").files.addUsingPath(fileNamePath, file, { Overwrite: true });
} else {
    // large upload
    result = await sp.web.getFolderByServerRelativePath("Shared Documents").files.addChunked(fileNamePath, file, 
        { progress: data => { setUploadProgress((data.offset/file.size)*100) }, 
          Overwrite: true 
        }
    );
}
console.log(result);
setIsUploading(false);
setIsDialogOpen(false);
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
          subText: 'Please fill in the details below to upload your file.'
        }}
      >
        <TextField label="File" type="file" onChange={handleFileChange} accept=".xlsx,.docx,.pptx,.zip" />
        <Dropdown
          label="Classification"
          options={classificationOptions}
          onChange={handleClassificationChange}
        />
        <TextField
          label="Validity (in days)"
          type="number"
          onChange={handleValidityChange}
        />
        {errorMessage && <div style={{ color: 'red' }}>{errorMessage}</div>}
        {isUploading && <ProgressIndicator label="Uploading..." percentComplete={uploadProgress / 100} />}
        <DialogFooter>
          <PrimaryButton onClick={handleUpload} text="Upload" />
          <DefaultButton onClick={() => setIsDialogOpen(false)} text="Cancel" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default MyDialogPopup;
