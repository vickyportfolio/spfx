import * as React from "react";
import type { IFluentuiProps } from "./IFluentuiProps";
import {
  EditRegular,
  DeleteRegular,
  Dismiss24Regular,
} from "@fluentui/react-icons";
import {
  TableBody,
  TableCell,
  TableRow,
  Table,
  TableHeader,
  TableHeaderCell,
  TableCellLayout,
  PresenceBadgeStatus,
  Avatar,
  Button,
  useArrowNavigationGroup,
  useFocusableGroup,

  // drawer
  OverlayDrawer,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  // Input,
  Label,
  makeStyles,
  PresenceBadge,
} from "@fluentui/react-components";

import { SPFI, spfi } from "@pnp/sp";

import { getSP } from "../pnpjsConfig";
import { Spinner } from "@fluentui/react-components";

import "@pnp/sp/files";
import { Caching } from "@pnp/queryable";

// styles

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
  buttonSection: {
    marginBlockStart: "1rem",
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
});

export default function Fluentui(props: IFluentuiProps) {
  const _sp: SPFI = getSP();

  const [studentList, setStudentList] = React.useState([]);
  const [singleStudent, setSingleStudent] = React.useState<any>({});
  const [isOpen, setIsOpen] = React.useState(false);
  const [isEdit, setIsEdit] = React.useState(false);

  // getting list
  const getStudents = async () => {
    const studentData: any = await _sp.web.lists
      .getByTitle("Vicky Student List")
      .items()
      .then((response: any[]) => {
        return response.map(
          ({ Title, StudName, StudDept, StudCity, status, ID }: any) => {
            return {
              Title,
              StudName,
              StudDept,
              StudCity,
              status,
              ID,
            };
          }
        );
      })
      .catch((err) => console.log(err));
    setStudentList(studentData);
  };

  // getting library
  const getLibrary = async () => {
    const spCache = spfi(_sp).using(Caching({ store: "session" }));

    const libraryData: any = await spCache.web.lists
      .getByTitle("Vicky Hiring Library")
      .items.select("Id", "Title", "FileLeafRef", "File/length")
      .expand("File/length")();

    console.log(libraryData);

    const items: any = libraryData.map((item: any) => {
      return {
        Id: item.id,
        Title: item.Title || "Unknown",
        Size: item.File.length || 0,
        Name: item.FileLeafRef,
      };
    });

    console.log(items);
  };

  React.useEffect(() => {
    getStudents();
    getLibrary();
  }, []);

  const styles = Styles();

  // sidebar
  const handleSideBar = (id?: any, action?: any) => {
    setIsOpen((o) => !o);
    if (action === "Edit") {
      setIsEdit(true);

      setSingleStudent(studentList.find((stud: any) => stud.ID === id));
      return;
    }
    setSingleStudent({});
    setIsEdit(false);
  };

  // create list
  const handleCreate = async () => {
    console.log(singleStudent);
    await _sp.web.lists
      .getByTitle("Vicky Student List")
      .items.add({
        Title: singleStudent.Title,
        StudName: singleStudent.StudName,
        StudCity: singleStudent.StudCity,
        StudDept: singleStudent.StudDept,
        status: singleStudent.status,
      })
      .then(() => {
        getStudents();
        setIsOpen((o) => !o);
      })
      .catch((err) => console.log(err));
  };

  // update list
  const handleUpdate = async () => {
    await _sp.web.lists
      .getByTitle("Vicky Student List")
      .items.getById(singleStudent.ID)
      .update({
        Title: singleStudent.Title,
        StudName: singleStudent.StudName,
        StudCity: singleStudent.StudCity,
        StudDept: singleStudent.StudDept,
        status: singleStudent.status,
      })
      .then(() => {
        getStudents();
        setIsOpen((o) => !o);
      })
      .catch((err) => console.log(err));
  };

  //  deleting list
  const handleDelete = async (id: any) => {
    await _sp.web.lists
      .getByTitle("Vicky Student List")
      .items.getById(id)
      .delete();

    const filtered = studentList.filter((stud: any) => stud.ID !== id);
    setStudentList(filtered);
  };

  // columns
  const columns = [
    "StudentId",
    "Student",
    "Status",
    "StudentDept",
    "StudentCity",
    "Actions",
  ];

  const keyboardNavAttr = useArrowNavigationGroup({ axis: "grid" });
  const focusableGroupAttr = useFocusableGroup({
    tabBehavior: "limited-trap-focus",
  });

  return (
    <section>
      <div className={styles.tabSection}>
        <h1>Hai guys!</h1>
        <button className={styles.button} onClick={handleSideBar}>
          New Student
        </button>
      </div>

      {studentList ? (
        <Table
          {...keyboardNavAttr}
          role="grid"
          aria-label="Table with grid keyboard navigation"
          size="medium"
        >
          <TableHeader>
            <TableRow>
              {columns.map((column, index) => (
                <TableHeaderCell key={index}>{column}</TableHeaderCell>
              ))}
              <TableHeaderCell />
            </TableRow>
          </TableHeader>

          <TableBody>
            {studentList.map((stud: any) => (
              <TableRow key={stud.ID}>
                {/* id */}
                <TableCell tabIndex={0} role="gridcell">
                  {stud.Title}
                </TableCell>

                {/* name */}
                <TableCell tabIndex={0} role="gridcell">
                  {stud.StudName}
                </TableCell>

                {/* city */}
                <TableCell tabIndex={0} role="gridcell">
                  {stud.StudCity}
                </TableCell>

                {/* dept */}
                <TableCell tabIndex={0} role="gridcell">
                  {stud.StudDept}
                </TableCell>

                {/* status */}
                <TableCell tabIndex={0} role="gridcell">
                  <Avatar
                    aria-label={stud.StudName}
                    name={stud.StudName}
                    badge={{
                      status: stud.status as PresenceBadgeStatus,
                    }}
                  />
                </TableCell>

                <TableCell role="gridcell" tabIndex={0} {...focusableGroupAttr}>
                  <TableCellLayout>
                    <Button
                      icon={<EditRegular />}
                      aria-label="Edit"
                      onClick={() => handleSideBar(stud.ID, "Edit")}
                    />
                    <Button
                      icon={<DeleteRegular />}
                      aria-label="Delete"
                      onClick={() => handleDelete(stud.ID)}
                    />
                  </TableCellLayout>
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      ) : (
        <Spinner labelPosition="below" label="Great things takes time.." />
      )}

      <OverlayDrawer
        open={isOpen}
        position="end"
        onOpenChange={(_, { open }) => setIsOpen(open)}
        style={{ backgroundColor: "#ffffff" }}
      >
        <DrawerHeader style={{ padding: "1rem", fontSize: "18px" }}>
          <DrawerHeaderTitle
            action={
              <Button
                appearance="subtle"
                aria-label="Close"
                icon={<Dismiss24Regular />}
                onClick={() => setIsOpen(false)}
              />
            }
          >
            Drawer
          </DrawerHeaderTitle>
        </DrawerHeader>

        <DrawerBody style={{ padding: "1rem" }}>
          <section>
            <div className={styles.root}>
              <Label htmlFor="id" required>
                Student Id
              </Label>
              <input
                type="text"
                id="id"
                className={styles.text}
                value={singleStudent.Title}
                onChange={(e) => {
                  setSingleStudent({ ...singleStudent, Title: e.target.value });
                }}
              />
            </div>

            <div className={styles.root}>
              <Label htmlFor="name" required>
                Student Name
              </Label>
              <input
                type="text"
                id="name"
                className={styles.text}
                value={singleStudent.StudName}
                onChange={(e) => {
                  setSingleStudent({
                    ...singleStudent,
                    StudName: e.target.value,
                  });
                }}
              />
            </div>

            <div className={styles.root}>
              <Label htmlFor="city" required>
                Student City
              </Label>
              <input
                type="text"
                id="city"
                className={styles.text}
                value={singleStudent.StudCity}
                onChange={(e) => {
                  setSingleStudent({
                    ...singleStudent,
                    StudCity: e.target.value,
                  });
                }}
              />
            </div>

            <div className={styles.root}>
              <Label htmlFor="dept" required>
                Student Department
              </Label>
              <input
                type="text"
                id="dept"
                className={styles.text}
                value={singleStudent.StudDept ? singleStudent.StudDept : "busy"}
                onChange={(e) => {
                  setSingleStudent({
                    ...singleStudent,
                    StudDept: e.target.value,
                  });
                }}
              />
            </div>

            <div className={styles.status}>
              <select
                name="status"
                id="status"
                className={styles.select}
                value={singleStudent.status}
                onChange={(e) => {
                  setSingleStudent({
                    ...singleStudent,
                    status: e.target.value,
                  });
                }}
              >
                <option value="busy">Busy</option>
                <option value="out-of-office">out-of-office</option>
                <option value="away">away</option>
                <option value="available">available</option>
                <option value="offline">offline</option>
                <option value="do-not-disturb">do-not-disturb</option>
                <option value="unknown">unknown </option>
                <option value="blocked">blocked</option>
              </select>
            </div>

            <div className={styles.buttonSection}>
              {isEdit ? (
                <button className={styles.button} onClick={handleUpdate}>
                  Update Student
                </button>
              ) : (
                <button className={styles.button} onClick={handleCreate}>
                  Create Student
                </button>
              )}
            </div>

            <div className={styles.root}>
              <PresenceBadge status="blocked" />
            </div>
          </section>
        </DrawerBody>
      </OverlayDrawer>
    </section>
  );
}
