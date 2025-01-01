import * as React from 'react';
import "@pnp/sp/context-info";
import { getSP } from "../pnpconfig";
import { spfi } from "@pnp/sp";
import { Caching } from "@pnp/queryable";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import Select, { SelectChangeEvent } from '@mui/material/Select';
import Box from '@mui/material/Box';
import LinearProgress, { LinearProgressProps } from '@mui/material/LinearProgress';
import Button from '@mui/material/Button';
import TextField from '@mui/material/TextField';
import Dialog from '@mui/material/Dialog';
import DialogActions from '@mui/material/DialogActions';
import DialogContent from '@mui/material/DialogContent';
import DialogTitle from '@mui/material/DialogTitle';
import FormControl from '@mui/material/FormControl';
import InputLabel from '@mui/material/InputLabel';
import MenuItem from '@mui/material/MenuItem';
import Dropzone from 'react-dropzone'
import Stack from '@mui/material/Stack';
import AddCircleOutlineIcon from '@mui/icons-material/AddCircleOutline';
import InfoOutlinedIcon from '@mui/icons-material/InfoOutlined';
import Typography from '@mui/material/Typography';
import FileOpenIcon from '@mui/icons-material/FileOpen';
import IconButton from '@mui/material/IconButton';
import DeleteOutlineIcon from '@mui/icons-material/DeleteOutline';





interface MyDialogPopupProps {
  absoluteURL: string
}
function LinearProgressWithLabel(props: LinearProgressProps & { value: number }) {
  return (
    <Box sx={{ display: 'flex', alignItems: 'center' }}>
      <Box sx={{ width: '100%', mr: 1 }}>
        <LinearProgress variant="determinate" {...props} />
      </Box>
      <Box sx={{ minWidth: 35 }}>
        <Typography
          variant="body2"
          sx={{ color: 'text.secondary' }}
        >{`${Math.round(props.value)}%`}</Typography>
      </Box>
    </Box>
  );
}


const MyDialogPopup: React.FC<MyDialogPopupProps> = ({ absoluteURL }) => {
  const [file, setFile] = React.useState<File | null>(null);
  const [classification, setClassification] = React.useState<string | undefined>(undefined);
  const [validity, setValidity] = React.useState<number | undefined>(undefined);
  const [errorMessage, setErrorMessage] = React.useState<string | null>(null);
  const [uploadProgress, setUploadProgress] = React.useState<number>(0);
  const [isUploading, setIsUploading] = React.useState<boolean>(false);
  const [open, setOpen] = React.useState<boolean>(false);
  const allowedFileTypes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'application/x-zip-compressed'
  ];
  const sp = spfi(getSP()).using(Caching({ store: "session" }));

  const handleClickOpen = () => {
    setOpen(true);
  };

  const handleClose = () => {
    setOpen(false);
  };

  const handleFile = (file: File) => {
    if (file && !allowedFileTypes.includes(file.type)) {
      setErrorMessage('Selected file type is not allowed');
      setFile(null);
    } else {
      setErrorMessage(null);
      setFile(file);
    }
  }



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
        {
          progress: data => { setUploadProgress((data.offset / file.size) * 100) },
          Overwrite: true
        }
      );
    }
    console.log(result);
    handleClose();
    setIsUploading(false);
  };

  return (
    <React.Fragment>
      <Button variant="outlined" onClick={handleClickOpen}>
        New File
      </Button>
      <Dialog
        open={open}
        onClose={handleClose}
        PaperProps={{
          component: 'form',
          onSubmit: (event: React.FormEvent<HTMLFormElement>) => {
            event.preventDefault();
            handleUpload();
            
          },
        }}
      >
        <DialogTitle>Upload new file</DialogTitle>
        <DialogContent style={{ paddingTop: 20 }}>
          <FormControl sx={{ mb: 2, minWidth: 120 }} size="small" fullWidth>
            <Dropzone onDrop={acceptedFiles => handleFile(acceptedFiles[0])}
            //accept={{ 'image/jpeg': [], 'image/png': [] }}
            >
              {({ getRootProps, getInputProps }) => (
                <Box component="section" sx={{ p: 2, border: '1px dashed grey' }} style={{ cursor: 'pointer' }}>
                  <div {...getRootProps()}>
                    <input {...getInputProps()} />
                    <Stack spacing={2} sx={{ justifyContent: "center", alignItems: "center" }}>
                      <AddCircleOutlineIcon sx={{ fontSize: 34 }} />
                      <Typography variant="body1">Drag & drop or click to select file</Typography>
                      <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                        <InfoOutlinedIcon sx={{ fontSize: 14 }} />
                        <Typography variant="body2" style={{ fontWeight: 'bold' }}>Max file size: 10 GB</Typography>
                      </div>
                    </Stack>
                  </div>
                </Box>
              )}
            </Dropzone>
          </FormControl>
          {file && (
            <FormControl sx={{ mb: 2, minWidth: 120 }} size="small" fullWidth>
              <Box component="section" sx={{ p: 2, border: '1px dashed grey' }}>
                <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
                  <FileOpenIcon color="secondary" />
                  <div style={{ flexGrow: 1 }}>
                    <Typography variant="body1" style={{ fontWeight: 'semiBold' }}>{file?.name}</Typography>
                    <Typography variant="body2" >{file ? (file.size / 1000000).toFixed(2) + ' MB' : ''}</Typography>
                  </div>
                  <IconButton onClick={() => { setFile(null) }}>
                    <DeleteOutlineIcon />
                  </IconButton>

                </div>
                {isUploading && (
                  <Box sx={{ width: '100%' }}>
                    <LinearProgressWithLabel value={uploadProgress} />

                  </Box>
                )}
              </Box>

            </FormControl>

          )}

          <FormControl sx={{ mb: 2, minWidth: 120 }} size="small" fullWidth>
            <InputLabel id="demo-simple-select-label">Classification</InputLabel>
            <Select
              labelId="demo-simple-select-label"
              id="demo-simple-select"
              value={classification}
              label="Classification"
              onChange={(event: SelectChangeEvent) => { setClassification(event.target.value as string) }}
            >
              <MenuItem value={10}>Public</MenuItem>
              <MenuItem value={20}>Private</MenuItem>
              <MenuItem value={30}>Confidential</MenuItem>
            </Select>
          </FormControl>
          <FormControl sx={{minWidth: 120 }} size="small" fullWidth>
            <TextField
              onChange={handleValidityChange}
              variant='outlined'
              id="outlined-number"
              label="Validity"
              size="small"
              type="number"

            />
          </FormControl>
          {errorMessage && (
            <h3>{errorMessage}</h3>
          )}
        </DialogContent>
        <DialogActions>
          <Button onClick={handleClose}>Cancel</Button>
          <Button type="submit">Upload</Button>
        </DialogActions>
      </Dialog>
    </React.Fragment>
  );
};

export default MyDialogPopup;
