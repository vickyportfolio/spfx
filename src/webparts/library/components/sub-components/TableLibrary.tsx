import * as React from "react";
import {
  EditRegular,
  DeleteRegular,
  DocumentRegular,
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
  Spinner,
  Button,
} from "@fluentui/react-components";

export const TableLibrary = ({
  library,
  handleSideBar,
  handleDelete,
  columns,
  Deleted,
}: any) => {
  const columnExists = columns.map(({ columnKey }: any) => columnKey);
  // console.log(columnExists);
  return (
    <section>
      {library.length > 0 ? (
        <Table arial-label="Default table">
          <TableHeader>
            <TableRow>
              {columns.map((column: any) => (
                <TableHeaderCell key={column.columnKey}>
                  <TableCellLayout media={column.icon}>
                    {column.label}
                  </TableCellLayout>
                </TableHeaderCell>
              ))}
            </TableRow>
          </TableHeader>

          <TableBody>
            {/* {columns.filter((col: any) => {})} */}
            {library.map((item: any) => (
              <TableRow key={item.id}>
                {/* document name */}
                {columnExists.includes("file") && (
                  <TableCell>
                    <TableCellLayout media={<DocumentRegular />}>
                      <a href={item.url} target="blank">
                        {item.file}
                      </a>
                    </TableCellLayout>
                  </TableCell>
                )}
                {/* Created By */}
                {columnExists.includes("creator") && (
                  <TableCell>
                    <TableCellLayout
                      truncate
                      media={
                        <Avatar
                          aria-label={item.createdBy}
                          name={item.createdBy}
                          badge={{
                            status: item.status as PresenceBadgeStatus,
                          }}
                        />
                      }
                    >
                      {item.createdBy}
                    </TableCellLayout>
                  </TableCell>
                )}

                {/* file size */}
                {columnExists.includes("FileSize") && (
                  <TableCell>
                    {(item.fileSize / 1024).toFixed(1) + "KB"}
                  </TableCell>
                )}

                {/* start date */}
                {columnExists.includes("StartDate") && (
                  <TableCell>{item.startDate}</TableCell>
                )}
                {/* end date */}
                {columnExists.includes("EndDate") && (
                  <TableCell>{item.endDate}</TableCell>
                )}
                {/* about */}
                {columnExists.includes("About") && (
                  <TableCell>{item.about}</TableCell>
                )}
                {/* approver/lookup */}
                {columnExists.includes("Approver") && (
                  <TableCell>{item.approver}</TableCell>
                )}
                {/* reviewer/choice */}
                {columnExists.includes("Reviewer") && (
                  <TableCell>{item.reviewer}</TableCell>
                )}
                {/* action */}
                {columnExists.includes("Action") && (
                  <TableCell role="gridcell" tabIndex={0}>
                    <TableCellLayout>
                      <Button
                        icon={<EditRegular />}
                        aria-label="Edit"
                        onClick={() => handleSideBar(item.id, "Edit")}
                      />
                      <Button
                        icon={<DeleteRegular />}
                        aria-label="Delete"
                        onClick={() => handleDelete(item.id)}
                      />
                    </TableCellLayout>
                  </TableCell>
                )}
              </TableRow>
            ))}
          </TableBody>
        </Table>
      ) : (
        <Spinner
          labelPosition="below"
          label={
            Deleted
              ? "No Records found"
              : "Grass is always greener on other side..."
          }
        />
      )}
    </section>
  );
};
