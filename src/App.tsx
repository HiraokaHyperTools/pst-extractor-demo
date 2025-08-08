import * as React from 'react';
import Container from '@mui/material/Container';
import Typography from '@mui/material/Typography';
import Box from '@mui/material/Box';
import { Consts, openPst, PSTAttachment, PSTFile, PSTFolder, PSTMessage } from '@hiraokahypertools/pst-extractor';
import AppBar from '@mui/material/AppBar';
import Toolbar from '@mui/material/Toolbar';
import Button from '@mui/material/Button';
import IconButton from '@mui/material/IconButton';
import ListSubheader from '@mui/material/ListSubheader';
import List from '@mui/material/List';
import ListItemButton from '@mui/material/ListItemButton';
import ListItemIcon from '@mui/material/ListItemIcon';
import ListItemText from '@mui/material/ListItemText';
import FolderIcon from '@mui/icons-material/Folder';
import MailIcon from '@mui/icons-material/Mail';
import PersonIcon from '@mui/icons-material/Person';
import QuestionMarkIcon from '@mui/icons-material/QuestionMark';
import Drawer from '@mui/material/Drawer';
import ListItem from '@mui/material/ListItem';
import Table from '@mui/material/Table';
import TableBody from '@mui/material/TableBody';
import TableCell from '@mui/material/TableCell';
import TableHead from '@mui/material/TableHead';
import TableRow from '@mui/material/TableRow';
import LinearProgress from '@mui/material/LinearProgress';
import ArrowBackIcon from '@mui/icons-material/ArrowBack';
import HomeIcon from '@mui/icons-material/Home';
import Card from '@mui/material/Card';
import CardContent from '@mui/material/CardContent';
import CardActions from '@mui/material/CardActions';
import AttachEmailIcon from '@mui/icons-material/AttachEmail';
import EventIcon from '@mui/icons-material/Event';


type PstLoader = {
  pstFile: PSTFile,
  name: string,
};
type FlattenFolder = {
  key: string,
  folder: PSTFolder,
  name: string,
  sub: string,
  depth: number,
}
type ContactSummary = {
  address: string,
  name: string,
};
type MessageSummary = {
  key: string,
  subject: string,
  messageClass: string,
  from: ContactSummary,
  message: PSTMessage,
};
type AttachmentSummary = {
  displayName: string,
  provideFile: (() => Promise<ArrayBuffer | null>) | null,
  provideEmbedded: (() => Promise<PSTMessage | null>) | null,
  toJSON: () => any,
};
type MessageDetail = {
  to: ContactSummary[],
  cc: ContactSummary[],
  bcc: ContactSummary[],
  attachments: AttachmentSummary[],
  body: string,
  moreProperties: Map<string, string>,
};
type ShowPropertiesPage = {
  moreProperties: Map<string, string>,
};
type NavigationStackItem = {
  displayName: string,
  onClick: () => void,
  onLeave: () => void,
};
type NextCommand = {
  backToAllFoldersList?: boolean,
  setSelectedFolder?: FlattenFolder,
  setSelectedMessage?: MessageSummary,
  performOpenPst?: boolean,
  performShowProperties?: { toJSON: () => any, displayName: string },
};

let justIndex = 0;

function TaskObserver(props: {
  task: PromiseLike<void> | null,
  idle: () => React.JSX.Element,
  loading: () => React.JSX.Element,
  error: (ex: Error) => React.JSX.Element,
  success: () => React.JSX.Element,
}) {
  const [elem, setElem] = React.useState<React.JSX.Element | null>(null);

  React.useEffect(() => {
    (async () => {
      if (props.task) {
        try {
          setElem(props.loading());
          await props.task;
          setElem(props.success());
        }
        catch (ex) {
          setElem(props.error(ex as Error));
        }
      }
    })();
  }, [props.task]);

  if (props.task) {
    return elem;
  }
  else {
    return props.idle();
  }
}

export default function App() {
  const [file, setFile] = React.useState<File | null>(null);
  const [openTask, setOpenTask] = React.useState<Promise<void> | null>(null);
  const [pstLoader, setPstLoader] = React.useState<PstLoader | null>(null);
  const [flattenFolders, setFlattenFolders] = React.useState<FlattenFolder[]>([]);
  const [selectedFolder, setSelectedFolder] = React.useState<FlattenFolder | null>(null);
  const [messages, setMessages] = React.useState<MessageSummary[] | null>(null);
  const [ansiEncoding, setAnsiEncoding] = React.useState<string>("windows-1252");
  const [open, setOpen] = React.useState(false);
  const [selectedMessage, setSelectedMessage] = React.useState<MessageSummary | null>(null);
  const [messageDetail, setMessageDetail] = React.useState<MessageDetail | null>(null);
  const [navigationStack, setNavigationStack] = React.useState<NavigationStackItem[]>([]);
  const [nextCommand, setNextCommand] = React.useState<NextCommand>({});
  const [showProperties, setShowProperties] = React.useState<ShowPropertiesPage | null>(null);

  function eject() {
    setFile(null);
    setOpenTask(null);
    setPstLoader(null);
    setFlattenFolders([]);
    setSelectedFolder(null);
    setMessages(null);
    setNavigationStack([]);
    justIndex = 0;
  }

  function closeMessageView() {
    setSelectedMessage(null);
    setMessageDetail(null);
  }
  function closeMessageList() {
    setMessages(null);
    setSelectedFolder(null);
    closeMessageView();
  }

  React.useEffect(() => {
    if (nextCommand) {
      if (nextCommand.backToAllFoldersList) {
        closeMessageList();
      }
      if (nextCommand.setSelectedFolder) {
        if (selectedFolder !== nextCommand.setSelectedFolder) {
          setSelectedFolder(nextCommand.setSelectedFolder);
        }
        else {
          closeMessageView();
        }
      }
      if (nextCommand.setSelectedMessage) {
        setSelectedMessage(nextCommand.setSelectedMessage);
      }
      if (nextCommand.performOpenPst) {
        setOpenTask((async () => {
          if (file) {
            const pstFile = await openPst(
              {
                readFile: async (buffer: ArrayBuffer, offset: number, length: number, position: number): Promise<number> => {
                  const blob = file.slice(position, position + length);
                  const arrayBuffer = await blob.arrayBuffer();
                  const srcArray = new Uint8Array(arrayBuffer);
                  const destArray = new Uint8Array(buffer);
                  destArray.set(srcArray, offset);
                  return destArray.byteLength;
                },
                close: async (): Promise<void> => {

                }
              },
              {
                ansiEncoding: ansiEncoding,
              }
            );
            setPstLoader({ pstFile: pstFile, name: file.name });
          }
          else {
            throw new Error("No file selected.");
          }
        })());
      }
      if (nextCommand.performShowProperties) {
        const { toJSON } = nextCommand.performShowProperties;
        const props = Object.entries(toJSON())
          .map(pair => [pair[0] + "", pair[1] + ""] as [string, string])
          .sort((a, b) => a[0].localeCompare(b[0]))
          ;
        setShowProperties({
          moreProperties: new Map<string, string>(props),
        });
        setNavigationStack(navigationStack.concat({
          displayName: nextCommand.performShowProperties.displayName,
          onClick: () => setNextCommand(nextCommand),
          onLeave: () => {
            setShowProperties(null);
          },
        }));
      }
    }
  }, [nextCommand]);

  function convertAttachmentToSummary(att: PSTAttachment): AttachmentSummary {
    return ({
      displayName: att.displayName || att.filename,
      provideFile: (false
        || att.attachMethod === PSTAttachment.ATTACHMENT_METHOD_BY_VALUE
        || att.attachMethod === PSTAttachment.ATTACHMENT_METHOD_BY_REFERENCE
      )
        ? async () => { return att.fileData || null; }
        : null,
      provideEmbedded: (false
        || att.attachMethod === PSTAttachment.ATTACHMENT_METHOD_EMBEDDED
      )
        ? async () => { return att.getEmbeddedPSTMessage() || null; }
        : null,
      toJSON: () => att.toJSON(),
    });
  }

  React.useEffect(() => {
    if (selectedMessage) {
      (async () => {
        const recipients = await selectedMessage.message.getRecipients();
        const props = Object.entries(selectedMessage.message.toJSON())
          .map(pair => [pair[0] + "", pair[1] + ""] as [string, string])
          .sort((a, b) => a[0].localeCompare(b[0]))
          ;
        setMessageDetail({
          to: recipients.filter(r => r.recipientType === Consts.MAPI_TO).map(r => composeContactSummaryFrom(r.addrType, r.emailAddress, r.displayName)),
          cc: recipients.filter(r => r.recipientType === Consts.MAPI_CC).map(r => composeContactSummaryFrom(r.addrType, r.emailAddress, r.displayName)),
          bcc: recipients.filter(r => r.recipientType === Consts.MAPI_BCC).map(r => composeContactSummaryFrom(r.addrType, r.emailAddress, r.displayName)),
          attachments: (await selectedMessage.message.getAttachments()).map(convertAttachmentToSummary),
          body: await selectedMessage.message.body,
          moreProperties: new Map<string, string>(props),
        });
      })();
    } else {
      setMessageDetail(null);
    }
  }, [selectedMessage]);

  React.useEffect(() => {
    (async () => {
      if (pstLoader) {
        const folder = await pstLoader.pstFile.getRootFolder();
        if (folder) {
          const flattenFolders: FlattenFolder[] = [];

          async function traverseFolder(folder: PSTFolder, depth: number) {
            flattenFolders.push({
              key: `_${justIndex++}`,
              folder,
              name: (folder.displayName.length)
                ? `${folder.displayName}`
                : "(Unnamed Folder)",
              sub: `${await folder.getEmailCount()} emails`,
              depth,
            });
            const subFolders = await folder.getSubFolders();
            for (const subFolder of subFolders) {
              await traverseFolder(subFolder, depth + 1);
            }
          }
          await traverseFolder(folder, 0);

          setFlattenFolders(flattenFolders);

          setNavigationStack(navigationStack.concat({
            displayName: pstLoader.name,
            onClick: () => setNextCommand({ backToAllFoldersList: true }),
            onLeave: () => {
              // no need to do anything here
            }
          }));
        }
      }
    })();
  }, [pstLoader]);

  function composeContactSummaryFrom(addrType: string, emailAddress: string, name: string): ContactSummary {
    return {
      address: `${addrType}: ${emailAddress}`,
      name: name,
    };
  }

  function summaryMessageFromPSTMessage(message: PSTMessage): MessageSummary {
    return {
      key: `_${justIndex++}`,
      subject: message.subject,
      messageClass: message.messageClass,
      from: composeContactSummaryFrom(message.senderAddrtype, message.senderEmailAddress, message.senderName),
      message: message,
    };
  }

  React.useEffect(() => {
    closeMessageView();

    if (selectedFolder) {
      (async () => {
        const folder = selectedFolder.folder;
        const emails = await folder.getEmails();
        setMessages(emails.map(summaryMessageFromPSTMessage));
      })();
    } else {
      setMessages(null);
    }
  }, [selectedFolder]);

  async function proceedDownload(provideFile: (() => Promise<ArrayBuffer | null>) | null, fileName: string) {
    if (provideFile) {
      const fileData = await provideFile();
      if (fileData) {
        const blob = new Blob([fileData], { type: "application/octet-stream" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }
    }
  }

  function browseFolder(folder: FlattenFolder) {
    setSelectedFolder(folder);

    setNavigationStack(navigationStack.concat({
      displayName: folder.name,
      onClick: () => setNextCommand({ setSelectedFolder: folder }),
      onLeave: () => {
        setSelectedFolder(null);
      }
    }));
  }

  function browseMessage(message: MessageSummary) {
    setSelectedMessage(message);

    setNavigationStack(navigationStack.concat({
      displayName: message.subject,
      onClick: () => setNextCommand({ setSelectedMessage: message }),
      onLeave: () => {
        closeMessageView();
      }
    }));
  }

  async function proceedView(provideEmbedded: (() => Promise<PSTMessage | null>) | null) {
    if (provideEmbedded) {
      const embeddedMessage = await provideEmbedded();
      if (embeddedMessage) {
        browseMessage(summaryMessageFromPSTMessage(embeddedMessage));
      }
    }
  }

  function tryToOpenPst() {
    setNextCommand({ performOpenPst: true });
  }

  function goBack() {
    if (2 <= navigationStack.length) {
      const leaving = navigationStack[navigationStack.length - 1];
      leaving.onLeave();
      const newStack = navigationStack.slice(0, -1);
      setNavigationStack(newStack);
      const lastItem = newStack[newStack.length - 1];
      lastItem.onClick();
    }
  }

  return pstLoader
    ? <>
      <Box sx={{ flexGrow: 1 }}>
        <AppBar position="static">
          <Toolbar>
            <IconButton
              size="large"
              edge="start"
              color="inherit"
              aria-label="menu"
              sx={{ mr: 2 }}
              onClick={() => goBack()}
            >
              {(2 <= navigationStack.length) ? <ArrowBackIcon /> : <HomeIcon />}
            </IconButton>

            <Typography variant="h6" component="div" sx={{ flexGrow: 1 }}>
              {(1 <= navigationStack.length) ? navigationStack[navigationStack.length - 1].displayName : "(Unmounted)"}
            </Typography>

            <Button color="inherit" onClick={() => eject()}>Eject</Button>
          </Toolbar>
        </AppBar>

        <Drawer open={open} onClose={() => setOpen(false)}>
          <Box sx={{ width: 250 }} role="presentation" onClick={() => setOpen(false)}>
            <List>
              <ListSubheader component="div" id="folder-menu-subheader">
                Navigation stack
              </ListSubheader>
              {
                navigationStack.map((it, index) => (
                  <ListItem key={index} disablePadding>
                    <ListItemButton onClick={() => {
                      it.onClick();
                      setNavigationStack(navigationStack.slice(0, index + 1));
                    }}>
                      <ListItemText primary={it.displayName} />
                    </ListItemButton>
                  </ListItem>
                ))
              }
            </List>
          </Box>
        </Drawer>

        {
          showProperties
            ? <>
              <Table size="small" aria-label="more-properties-table">
                <TableHead>
                  <TableRow>
                    <TableCell>Property</TableCell>
                    <TableCell>Value</TableCell>
                  </TableRow>
                </TableHead>
                <TableBody>
                  {
                    Array.from(showProperties.moreProperties.entries()).map(([key, value]) => (
                      <TableRow key={key}>
                        <TableCell component="th" scope="row">{key}</TableCell>
                        <TableCell>{value}</TableCell>
                      </TableRow>
                    ))
                  }
                </TableBody>
              </Table>
            </>
            : selectedMessage
              ? messageDetail
                ? <>
                  <Box sx={{ px: 1 }}>
                    <Button variant="text" onClick={() => setNextCommand({
                      performShowProperties:
                      {
                        toJSON: () => selectedMessage.message.toJSON(),
                        displayName: `Properties of message: ${selectedMessage.message.subject}`
                      }
                    })}>Show message's full properties</Button>
                  </Box>

                  <Table aria-label="message details table" size="small">
                    <TableHead>
                      <TableRow>
                        <TableCell>Message header</TableCell>
                        <TableCell>Contents</TableCell>
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {[
                        { name: "To", list: messageDetail.to },
                        { name: "Cc", list: messageDetail.cc },
                        { name: "Bcc", list: messageDetail.bcc }
                      ].map((row) => (
                        <TableRow
                          key={row.name}
                          sx={{ '&:last-child td, &:last-child th': { border: 0 } }}
                        >
                          <TableCell component="th" scope="row">{row.name}</TableCell>
                          <TableCell>{row.list.map(
                            one => <p><i>{one.name} &lt;{one.address}&gt;</i></p>
                          )}</TableCell>
                        </TableRow>
                      ))}
                      <TableRow
                        key="subject"
                        sx={{ '&:last-child td, &:last-child th': { border: 0 } }}
                      >
                        <TableCell component="th" scope="row">Subject</TableCell>
                        <TableCell>{selectedMessage.subject}</TableCell>
                      </TableRow>

                      <TableRow
                        key="attachments"
                        sx={{ '&:last-child td, &:last-child th': { border: 0 } }}
                      >
                        <TableCell component="th" scope="row">Attachments</TableCell>
                        <TableCell>
                          <Box
                            sx={{
                              width: '100%',
                              display: 'grid',
                              gridTemplateColumns: 'repeat(auto-fill, minmax(min(200px, 100%), 1fr))',
                              gap: 2,
                            }}
                          >
                            {messageDetail.attachments.map(
                              att => <Card variant="outlined" key={att.displayName}>
                                <CardContent>
                                  <Typography gutterBottom sx={{ color: 'text.secondary', fontSize: 14 }}>
                                    {att.displayName}
                                  </Typography>
                                </CardContent>
                                <CardActions>
                                  {att.provideFile
                                    ? <Button onClick={() => proceedDownload(att.provideFile, att.displayName)} size="small">Download</Button>
                                    : att.provideEmbedded
                                      ? <Button onClick={() => proceedView(att.provideEmbedded)} size="small">View</Button>
                                      : <span>(Unknown)</span>
                                  }
                                  <Button onClick={() => setNextCommand({
                                    performShowProperties:
                                    {
                                      toJSON: () => att.toJSON(),
                                      displayName: `Properties of attachment: ${att.displayName}`
                                    }
                                  })} color="secondary" size="small">Properties</Button>
                                </CardActions>
                              </Card>
                            )}
                          </Box>
                        </TableCell>
                      </TableRow>
                    </TableBody>
                  </Table>

                  <Box sx={{ p: 2 }}>
                    <Typography variant="h6">Body</Typography>
                    <div>{messageDetail.body}</div>
                  </Box>
                </>
                : <>
                  <Typography variant="body1" gutterBottom>
                    Loading message details in progress ...
                  </Typography>
                  <Box sx={{ width: '100%' }}>
                    <LinearProgress />
                  </Box>

                </>
              :
              selectedFolder
                ? messages
                  ?
                  <>
                    <Box sx={{ px: 1 }}>
                      <Button variant="text" onClick={() => setNextCommand({
                        performShowProperties:
                        {
                          toJSON: () => selectedFolder.folder.toJSON(),
                          displayName: `Properties of folder: ${selectedFolder.name}`
                        }
                      })}>Show folder's full properties</Button>
                    </Box>
                    <List
                      sx={{ width: '100%', bgcolor: 'background.paper' }}
                      component="nav"
                      aria-labelledby="messages-subheader"
                      subheader={
                        <ListSubheader component="div" id="messages-subheader">
                          Messages list
                        </ListSubheader>
                      }
                    >
                      {
                        messages.map(
                          it => <ListItemButton
                            key={it.key}
                            onClick={() => {
                              browseMessage(it);
                            }}
                          >
                            {
                              it.messageClass === "IPM.Note"
                                ? <ListItemIcon>
                                  <MailIcon />
                                </ListItemIcon>
                                : it.messageClass === "IPM.Contact"
                                  ? <ListItemIcon>
                                    <PersonIcon />
                                  </ListItemIcon>
                                  : it.messageClass === "IPM.Appointment"
                                    ? <ListItemIcon>
                                      <EventIcon />
                                    </ListItemIcon>
                                    : it.messageClass.indexOf("IPM.Document") === 0
                                      // `IPM.Document.jpegfile`, `IPM.Document.AcroExch.Document`, `IPM.Document.Word.Document.12` or similar
                                      ? <ListItemIcon>
                                        <AttachEmailIcon />
                                      </ListItemIcon>
                                      : <ListItemIcon>
                                        <QuestionMarkIcon />
                                      </ListItemIcon>
                            }
                            <ListItemText primary={it.subject} />
                          </ListItemButton>
                        )
                      }
                    </List>
                  </>
                  : <>
                    <Typography variant="body1" gutterBottom>
                      Loading emails in progress ...
                    </Typography>
                    <Box sx={{ width: '100%' }}>
                      <LinearProgress />
                    </Box>

                  </>
                : flattenFolders.length
                  ? <>
                    <List
                      sx={{ width: '100%', bgcolor: 'background.paper' }}
                      component="nav"
                      aria-labelledby="all-folders-subheader"
                      subheader={
                        <ListSubheader component="div" id="all-folders-subheader">
                          All folders list
                        </ListSubheader>
                      }
                    >
                      {
                        flattenFolders.map(
                          it => <ListItemButton
                            key={it.key}
                            sx={{ pl: 2 + it.depth * 2 }}
                            onClick={() => {
                              browseFolder(it);
                            }} >
                            <ListItemIcon>
                              <FolderIcon />
                            </ListItemIcon>
                            <ListItemText primary={it.name} secondary={it.sub} />
                          </ListItemButton>
                        )
                      }
                    </List>
                  </>
                  : <>
                    <Typography variant="body1" gutterBottom>
                      Loading folder list in progress ...
                    </Typography>
                    <Box sx={{ width: '100%' }}>
                      <LinearProgress />
                    </Box>
                  </>
        }
      </Box >
    </>
    : <>
      <Container maxWidth="sm">
        <Box sx={{ my: 4 }}>
          <p>Select an Outlook .pst file to open:</p>
          <p>
            <input type="file" accept=".pst" onChange={async e => {
              for (let it of e.target.files ?? []) {
                setFile(it);
                break;
              }
            }} />
          </p>
          <p>ANSI Encoding (e.g., windows-1252)</p>
          <p>
            <input type="text" value={ansiEncoding} onChange={e => setAnsiEncoding(e.target.value)} />
          </p>
          <p>
            <TaskObserver
              task={openTask}
              idle={() =>
                <button onClick={() => tryToOpenPst()}>Open</button>
              }
              loading={() =>
                <>
                  <button disabled>Open</button>
                  &nbsp;
                  <i>(Load in progress ...)</i>
                </>
              }
              error={ex => <>
                <pre>{ex.message}</pre>
                <button onClick={() => tryToOpenPst()}>Open</button>
              </>}
              success={() =>
                <>File opened successfully</>
              }
            />
          </p>
        </Box>
      </Container >
    </>
    ;
}
