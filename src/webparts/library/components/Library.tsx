import * as React from "react";
import { getSP } from "../pnpConfig";
import { TableLibrary } from "./sub-components/TableLibrary";
import { DocumentRegular } from "@fluentui/react-icons";

import "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/fields";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import { SPFI } from "@pnp/sp";
import { getGUID } from "@pnp/core";
import {
  OverlayDrawer,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  Button,
  makeStyles,
  Label,
  Field,
} from "@fluentui/react-components";
import { Dismiss24Regular } from "@fluentui/react-icons";
import { DatePicker } from "@fluentui/react-datepicker-compat";

const Styles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    rowGap: "5px",
    maxWidth: "300px",
    paddingBlockEnd: "1rem",
  },
  text: {
    padding: "0.5rem",
  },
  button: {
    padding: "0.5rem",
    cursor: "pointer",
    fontSize: "bold",
  },
  tabSection: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  control: {
    maxWidth: "300px",
  },
  buttonSection: {
    marginBlockStart: "1rem",
  },
  status: {
    display: "grid",
    gridTemplateRows: "repeat(1fr)",
    justifyItems: "start",
    gap: "2px",
    maxWidth: "400px",
    paddingBlockStart: "1rem",
  },
  select: {
    paddingInline: "1rem",
    width: "200px",
    paddingBlock: "0.5rem",
  },
  disabled: {
    cursor: "not-allowed",
    pointerEvents: "none",
  },
  leftPanel: {
    display: "flex",
    alignItems: "center",
    columnGap: "1rem",
  },
});

function Library() {
  const _sp: SPFI = getSP();
  const AllCOLUMNS = [
    { columnKey: "file", label: "File", icon: <DocumentRegular /> },
    { columnKey: "creator", label: "Created By" },
    { columnKey: "FileSize", label: "File size" },
    { columnKey: "StartDate", label: "Start Date" },
    { columnKey: "EndDate", label: "End Date" },
    { columnKey: "About", label: "About" },
    { columnKey: "Action", label: "Action" },
  ];
  const [library, setLibrary] = React.useState<any>([]);
  const [columns, setColumns] = React.useState<any>(AllCOLUMNS);
  const [isOpen, setIsOpen] = React.useState(false);
  const [documentFile, setDocument] = React.useState<any>();
  const [isEdit, setIsEdit] = React.useState(false);
  const [singleFile, setSingleFile] = React.useState<any>({});
  const [startvalue, setStartValue] = React.useState<Date | null | undefined>(
    null
  );
  const [endvalue, setEndValue] = React.useState<Date | null | undefined>(null);
  const datePickerRefEnd = React.useRef<HTMLInputElement>(null);
  const datePickerRefStart = React.useRef<HTMLInputElement>(null);
  const [isDeleted, setisDeleted] = React.useState<Boolean>(false);
  const [Approver, setApprover] = React.useState<any>([]);

  const Style = Styles();

  React.useEffect(() => {
    getLibrary();
  }, []);

  // geting/read library details
  async function getLibrary() {
    const LibraryDetails = await _sp.web
      .getFolderByServerRelativePath("Vicky hiring Library")
      .files.expand("ListItemAllFields/FieldValuesAsText")()
      .then((res: any[]) => {
        return res.map((item: any, index: any) => {
          /* const startDate = `${new Date(
            item.ListItemAllFields.StartDate
          ).getUTCDate()}/${
            new Date(item.ListItemAllFields.StartDate).getUTCMonth() + 1
          }/${
            new Date(item.ListItemAllFields.StartDate).getUTCFullYear() % 100
          }`;

          const endDate = `${new Date(
            item.ListItemAllFields.EndDate
          ).getUTCDate()}/${
            new Date(item.ListItemAllFields.EndDate).getUTCMonth() + 1
          }/${new Date(item.ListItemAllFields.EndDate).getUTCFullYear() % 100}`; */

          return {
            file: item.Name,
            url: item.LinkingUrl,
            fileSize: item.Length,
            startDate: item.ListItemAllFields.FieldValuesAsText.StartDate,
            endDate: item.ListItemAllFields.FieldValuesAsText.EndDate,
            id: item.ListItemAllFields.ID,
            redirectUrl: item.ListItemAllFields.ServerRedirectedEmbedUrl,
            about:
              item.ListItemAllFields.FieldValuesAsText
                .Hiring_x005f_x0020_x005f_Description,
            status: item.ListItemAllFields.status,
            approver: item.ListItemAllFields.FieldValuesAsText.Approver,
            reviewer: item.ListItemAllFields.reviewers,
            createdBy: item.ListItemAllFields.FieldValuesAsText.Author,
            documentName: item.Name,
          };
        });
      });

    console.log("Library", LibraryDetails);

    setLibrary(LibraryDetails);

    const approver = await _sp.web.lists
      .getByTitle("Vicky Student List")
      .items();
    setApprover(approver);

    // testing purpose calls
    /* const All_ITEMS = await _sp.web
      .getFolderByServerRelativePath("Vicky hiring Library")
      .files.expand("ListItemAllFields/FieldValuesAsText")();

    console.log("Allitems", All_ITEMS);

    const dummyList = await _sp.web.lists
      .getByTitle("Vicky hiring Library")
      .items();
    console.log("list", dummyList);

    const lookupVal = await _sp.web.lists
      .getByTitle("Vicky hiring Library")
      .items.select(
        "Approver/Title,Approver/Id,Approver/StudName,Approver/StudCity"
      )
      .expand("Approver")();

    console.log("Lookup =>", lookupVal);

    const item = await _sp.web.lists
      .getByTitle("Vicky hiring Library")
      .items.getById(55)
      .select("Id", "File/length")
      .expand("File/length")();

    console.log("item =>", item); */
  }

  // create
  const handleCreate = async () => {
    if (!documentFile) {
      return;
    }

    await _sp.web
      .getFolderByServerRelativePath("Vicky hiring Library")
      .files.addUsingPath(documentFile.name, documentFile, {
        Overwrite: true,
      })
      .then(async (item: any) => {
        const Filelist: any = await _sp.web
          .getFileByServerRelativePath(item.ServerRelativeUrl)
          .getItem();

        let userStatus: string;
        if (!singleFile.status) {
          userStatus = "busy";
        } else {
          userStatus = singleFile.status;
        }
        if (!singleFile.ApproverId) {
          singleFile.ApproverId = "9";
        }

        await Filelist.update({
          Hiring_x0020_Description: singleFile.about,
          status: userStatus,
          reviewers: singleFile.reviewer,
          StartDate: startvalue,
          EndDate: endvalue,
          ApproverId: singleFile.ApproverId,
        });

        if (singleFile.documentName) {
          const item = await _sp.web.lists
            .getByTitle("Vicky hiring Library")
            .items.getById(Filelist.ID)
            .select("Id", "Title", "FileLeafRef", "File/length")
            .expand("File/length");

          await item.update({ FileLeafRef: singleFile.documentName });
        }
      });
    console.log(singleFile);

    await getLibrary();
    setIsOpen((o) => !o);
  };

  // update
  const handleUpdate = async () => {
    if (documentFile) {
      await _sp.web
        .getFolderByServerRelativePath("Vicky hiring Library")
        .files.getByUrl(singleFile.file)
        .setContent(documentFile);
    }
    let userStatus: string;
    if (!singleFile.status) {
      userStatus = "busy";
    } else {
      userStatus = singleFile.status;
    }
    let updationList = library.find((lib: any) => lib.id === singleFile.id);
    let libindex = library.findIndex((lib: any) => lib.id === singleFile.id);
    updationList.about = singleFile.about;
    updationList.status = singleFile.status;
    updationList.approver = singleFile.approver;
    updationList.reviewer = singleFile.reviewer;

    const updateList = await _sp.web.lists
      .getByTitle("Vicky hiring Library")
      .items.getById(singleFile.id);

    await updateList.update({
      Hiring_x0020_Description: singleFile.about,
      status: userStatus,
      reviewers: singleFile.reviewer,
    });

    if (typeof startvalue !== "string") {
      await updateList.update({
        StartDate: startvalue,
      });
      const start: any = startvalue;
      const startTransform = new Date(start);
      const finalStart =
        startTransform.getMonth() +
        1 +
        "/" +
        startTransform.getDate() +
        "/" +
        startTransform.getFullYear();

      updationList.startDate = finalStart;
    }
    if (typeof endvalue !== "string") {
      await updateList.update({
        EndDate: endvalue,
      });
      const end: any = endvalue;
      const finalEnd =
        end.getMonth() + 1 + "/" + end.getDate() + "/" + end.getFullYear();
      updationList.endDate = finalEnd;
    }

    if (singleFile.documentName) {
      const item = await _sp.web.lists
        .getByTitle("Vicky hiring Library")
        .items.getById(singleFile.id)
        .select("Id", "Title", "FileLeafRef", "File/length")
        .expand("File/length");

      await item.update({ FileLeafRef: singleFile.documentName });
      updationList.file = singleFile.documentName + "xlsx";
    }

    if (singleFile.approver) {
      await _sp.web.lists
        .getByTitle("Vicky hiring Library")
        .items.getById(singleFile.id)
        .update({
          ApproverId: singleFile.ApproverId,
        });
    }
    let newLib = library;
    newLib[libindex] = updationList;

    setLibrary(newLib);
    setIsOpen((o) => !o);
  };

  // Delete
  const handleDelete = async (id?: any) => {
    await _sp.web.lists
      .getByTitle("Vicky hiring Library")
      .items.getById(id)
      .delete()
      .then((_): any => {
        const filtered = library.filter((stud: any) => stud.id !== id);
        if (filtered.length > 0) {
          setLibrary(filtered);
        } else {
          setisDeleted(true);
          setLibrary([]);
        }
      });
  };

  // upload
  const handleUpload = async (event: any) => {
    const uploadedFile = event.target.files[0];
    let showfile = document.getElementById("showFile") as HTMLElement;
    showfile.textContent = uploadedFile.name;
    setDocument(uploadedFile);
  };

  // sidebar
  const handleSideBar = (id?: any, action?: any) => {
    setIsOpen((o) => !o);
    if (action === "Edit") {
      const singleRecord = library.find((file: any) => file.id == id);
      setSingleFile(singleRecord);
      setStartValue(singleRecord.startDate);
      setEndValue(singleRecord.endDate);
      setIsEdit(true);
      return;
    }
    setSingleFile({});
    setIsEdit(false);
    setStartValue(null);
    setEndValue(null);
  };

  // clear start datepicker
  const onClickStart = React.useCallback((): void => {
    setStartValue(null);
    console.log(singleFile);
    // setSingleFile({ ...singleFile, startDate: null });
    datePickerRefStart.current?.removeAttribute("disabled");
    datePickerRefStart.current?.focus();
  }, []);

  // clear end datepicker
  const onClickEnd = React.useCallback((singleFile: any): void => {
    setEndValue(null);
    console.log(singleFile);
    datePickerRefEnd.current?.removeAttribute("disabled");

    datePickerRefEnd.current?.focus();
  }, []);

  const onFormatDate = (date?: any): string => {
    if (typeof date == "string") {
      return date;
    } else if (!date) {
      return "";
    } else {
      const assignedDate =
        date.getDate() + "/" + (date.getMonth() + 1) + "/" + date.getFullYear();

      return assignedDate;
    }
  };

  const onParseDateFromString = React.useCallback(
    (newValue: string): Date => {
      const previousValue = endvalue || new Date();
      const newValueParts = (newValue || "").trim().split("/");
      const day =
        newValueParts.length > 0
          ? Math.max(1, Math.min(31, parseInt(newValueParts[0], 10)))
          : previousValue.getDate();
      const month =
        newValueParts.length > 1
          ? Math.max(1, Math.min(12, parseInt(newValueParts[1], 10))) - 1
          : previousValue.getMonth();
      let year =
        newValueParts.length > 2
          ? parseInt(newValueParts[2], 10)
          : previousValue.getFullYear();
      if (year < 100) {
        year +=
          previousValue.getFullYear() - (previousValue.getFullYear() % 100);
      }
      return new Date(year, month, day);
    },
    [endvalue]
  );
  // view change
  const handleChange = (event: any) => {
    if (event.target.value === "custom") {
      let custom = columns.filter((column: any) => {
        if (
          column.columnKey !== "FileSize" &&
          column.columnKey !== "Action" &&
          column.columnKey !== "About"
        ) {
          return column;
        }
      });
      custom.push(
        { columnKey: "Approver", label: "Approver" },
        { columnKey: "Reviewer", label: "Reviewer" }
      );

      setColumns(custom);
    } else {
      setColumns(AllCOLUMNS);
    }
  };

  return (
    <section>
      <div className={Style.tabSection}>
        <h1>Library details</h1>

        <div className={Style.leftPanel}>
          <select
            name="view"
            id="view"
            className={Style.select}
            onChange={(e) => handleChange(e)}
          >
            <option value="All">All</option>
            <option value="custom">Custom view</option>
          </select>

          <button className={Style.button} onClick={handleSideBar}>
            New File
          </button>
        </div>
      </div>

      <TableLibrary
        library={library}
        columns={columns}
        handleSideBar={handleSideBar}
        handleDelete={handleDelete}
        Deleted={isDeleted}
      />

      <OverlayDrawer
        open={isOpen}
        position="end"
        onOpenChange={(_, { open }) => {
          setIsOpen(open);
        }}
        style={{ backgroundColor: "#ffffff" }}
      >
        <DrawerHeader style={{ padding: "1rem", fontSize: "18px" }}>
          <DrawerHeaderTitle
            action={
              <Button
                appearance="subtle"
                aria-label="Close"
                icon={<Dismiss24Regular />}
                onClick={() => {
                  setIsOpen(false);
                  setDocument(null);
                }}
              />
            }
          >
            <h1>Document Library</h1>
          </DrawerHeaderTitle>
        </DrawerHeader>

        <DrawerBody style={{ padding: "1rem" }}>
          <section>
            {/* file name */}
            <div className={Style.root}>
              <Label htmlFor="name">File Name</Label>
              <input
                type="text"
                id="name"
                className={Style.text}
                value={singleFile.documentName}
                onChange={(e) => {
                  setSingleFile({
                    ...singleFile,
                    documentName: e.target.value,
                  });
                }}
              />
            </div>
            {/* about file */}
            <div className={Style.root}>
              <Field label="About File">
                <textarea
                  value={singleFile.about}
                  onChange={(e) => {
                    setSingleFile({ ...singleFile, about: e.target.value });
                  }}
                />
              </Field>
            </div>
            {/* startdate */}
            <div className={Style.root}>
              <Field label="Start Date" required>
                <DatePicker
                  ref={datePickerRefStart}
                  allowTextInput
                  value={startvalue}
                  className={Style.control}
                  onSelectDate={setStartValue as (date?: Date | null) => void}
                  formatDate={onFormatDate}
                  parseDateFromString={onParseDateFromString}
                  placeholder="Select a starting date..."
                  disabled={singleFile.startDate ? true : false}
                />
              </Field>
              <button onClick={onClickStart} className={Style.button}>
                Clear
              </button>
            </div>
            {/* enddate */}
            <div className={Style.root}>
              <Field label="End Date" required>
                <DatePicker
                  ref={datePickerRefEnd}
                  allowTextInput
                  value={endvalue}
                  onSelectDate={setEndValue as (date?: Date | null) => void}
                  formatDate={onFormatDate}
                  parseDateFromString={onParseDateFromString}
                  placeholder="Select a ending date..."
                  className={Style.control}
                  disabled={singleFile.endDate ? true : false}
                />
              </Field>
              <button
                onClick={() => onClickEnd(singleFile)}
                className={Style.button}
              >
                Clear
              </button>
            </div>
            {/* file upload */}
            <div className={Style.root}>
              <input
                type="file"
                id="fileInput"
                onChange={(e) => handleUpload(e)}
                required
              />
              <pre id="showFile">
                {isEdit ? singleFile.file : documentFile?.name}
              </pre>
            </div>
            {/* approver  */}
            <div className={Style.root}>
              <Label htmlFor="approver" required>
                Approver
              </Label>
              <select
                name="approver"
                className={Style.select}
                value={singleFile.approver}
                onChange={(e: any) => {
                  setSingleFile({
                    ...singleFile,
                    ApproverId:
                      e.target[e.target.selectedIndex].getAttribute("data-id"),
                    approver: e.target.value,
                  });
                }}
              >
                {Approver.map((look: any) => {
                  return (
                    <option
                      value={look.StudName}
                      data-id={look.ID}
                      key={look.ID}
                    >
                      {look.StudName}
                    </option>
                  );
                })}
              </select>
            </div>
            {/* reviewer */}
            <div className={Style.root}>
              <Label htmlFor="reviewer" required>
                Reviewer
              </Label>
              <input
                type="text"
                id="name"
                className={Style.text}
                value={singleFile.reviewer}
                onChange={(e) => {
                  setSingleFile({ ...singleFile, reviewer: e.target.value });
                }}
              />
            </div>
            {/* status */}
            <div className={Style.status}>
              <Field label="Status" required>
                <select
                  name="status"
                  id="status"
                  className={Style.select}
                  value={singleFile.status}
                  onChange={(e) => {
                    setSingleFile({
                      ...singleFile,
                      status: e.target.value,
                    });
                  }}
                >
                  <option value="busy" selected>
                    busy
                  </option>
                  <option value="out-of-office">out-of-office</option>
                  <option value="away">away</option>
                  <option value="available">available</option>
                  <option value="offline">offline</option>
                  <option value="do-not-disturb">do-not-disturb</option>
                  <option value="unknown">unknown </option>
                  <option value="blocked">blocked</option>
                </select>
              </Field>
            </div>
            {/* create/update */}
            <div className={Style.buttonSection}>
              {isEdit ? (
                <button className={Style.button} onClick={handleUpdate}>
                  Update Record
                </button>
              ) : (
                <button className={Style.button} onClick={handleCreate}>
                  Create Record
                </button>
              )}
            </div>
          </section>
        </DrawerBody>
      </OverlayDrawer>
    </section>
  );
}

export default Library;
