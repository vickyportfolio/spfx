import * as React from "react";
import "./HelloWorld.scss";
import type { IHelloWorldProps } from "./IHelloWorldProps";
import { escape } from "@microsoft/sp-lodash-subset";
import * as pnp from "sp-pnp-js";

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  state = {
    studentList: [],
  };

  // after ui renders
  componentDidMount(): void {
    this._getListData();
  }

  // main render
  public render(): React.ReactElement<IHelloWorldProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      userDisplayName,
      para,
      checkBox,
      dropDown,
      toggle,
      domElement,
    } = this.props;

    return (
      <section className="helloWorld">
        <div className="welcome">
          <img
            className="welcomeImage"
            alt=""
            src={
              isDarkTheme
                ? require("../assets/welcome-dark.png")
                : require("../assets/welcome-light.png")
            }
          />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>

          <div className="webpart_property">
            <h2>Web part property values:</h2>

            <p>
              <strong>Description :</strong>
              {escape(description)}
            </p>
            <p>
              <strong>Multiline description :</strong>
              {escape(para)}
            </p>
            <p>
              <strong>Checked :</strong>
              {checkBox ? "checked" : "unchecked"}
            </p>
            <p>
              <strong>Dropdown :</strong>
              {dropDown}
            </p>
            <p>
              <strong>toggle :</strong>
              {toggle ? "on" : "off"}
            </p>
          </div>

          {this._renderList(domElement)}
        </div>
      </section>
    );
  }

  // rendering table
  private _renderList(domElement: any) {
    return (
      <section className={"crudSection"}>
        <table>
          <tr>
            <td>Student id</td>
            <td>
              <input type="text" id="studentId" tabIndex={1} />{" "}
            </td>
            <td className="btn-group">
              <input
                type="submit"
                value="Get Details"
                id="btnGet"
                tabIndex={-1}
                onClick={() => this._getListById(domElement)}
              />
              <input
                type="submit"
                value="Reset"
                id="btnReset"
                // tabIndex={-2}
                onClick={() => this._getListData()}
              />
            </td>
          </tr>

          <tr>
            <td>Student Name</td>
            <td>
              <input type="text" id="txtStudentName" tabIndex={2} />
            </td>
          </tr>
          <tr>
            <td>Student department</td>
            <td>
              <input type="text" id="txtStudentDept" tabIndex={3} />
            </td>
          </tr>
          <tr>
            <td>Student City</td>
            <td>
              <input type="text" id="txtStudentCity" tabIndex={4} />
            </td>
          </tr>
          <tr>
            <td colSpan={4}>
              <input
                type="submit"
                value="Insert"
                id="btnInsert"
                tabIndex={5}
                onClick={() => this._insertStudent(domElement)}
              />
            </td>
          </tr>
        </table>

        <div id="studStatus">
          <div className="stud_container">
            {this.state.studentList.length > 0 ? (
              this.state.studentList.map((stud: any) => {
                return (
                  <div className="stud_card" key={stud.ID}>
                    <div className="stud_card_details">
                      <label>
                        <h2
                          className={`stud_name${stud.ID}`}
                          spellCheck={true}
                          contentEditable
                        >
                          {stud.StudName}
                        </h2>
                      </label>

                      <label>
                        ID :{" "}
                        <span
                          className={`stud_id${stud.ID}`}
                          spellCheck={true}
                          contentEditable
                        >
                          {stud.Title}
                        </span>
                      </label>

                      <label>
                        Department :
                        <span
                          className={`stud_dept${stud.ID}`}
                          spellCheck={true}
                          contentEditable
                        >
                          {stud.StudDept}
                        </span>
                      </label>

                      <label>
                        Address :{" "}
                        <span
                          className={`stud_city${stud.ID}`}
                          spellCheck={true}
                          contentEditable
                        >
                          {stud.StudCity}
                        </span>
                      </label>
                    </div>

                    <div className="btn_group">
                      <input
                        type="submit"
                        value="Update"
                        id="btnUpdate"
                        onClick={() => this._updateStudent(domElement, stud.ID)}
                      />
                      <input
                        type="submit"
                        value="Delete"
                        id="btnDelete"
                        onClick={() => this._deleteStudent(stud.ID)}
                      />
                    </div>
                  </div>
                );
              })
            ) : (
              <div className="loading">loading...</div>
            )}
          </div>
        </div>
      </section>
    );
  }

  // student crud events
  private async _insertStudent(domElement: any) {
    let studentId = domElement.querySelector("#studentId")?.["value"];
    let studentName = domElement.querySelector("#txtStudentName")?.value;
    let studentDept = domElement.querySelector("#txtStudentDept")?.value;
    let studentCity = domElement.querySelector("#txtStudentCity")?.value;

    // checking already exists!
    const isExist = this.state.studentList.find((stud: any) => {
      return Number(stud.Title) === Number(studentId);
    });

    if (isExist) {
      alert(
        `Student with this ${studentId} is already registered! select different Id`
      );
      domElement.querySelector("#studentId").focus();
      return;
    }
    const data = {
      Title: studentId,
      StudName: studentName,
      StudDept: studentDept,
      StudCity: studentCity,
    };
    await pnp.sp.web.lists
      .getByTitle("Vicky Student List")
      .items.add(data)
      .then((response) => {
        alert("Students has been added successfully!");
        console.log(response);
        this.setState((prev: any) => ({
          studentList: [...prev.studentList, data],
        }));
      })
      .catch((err) => console.log(err));

    domElement.querySelector("#studentId").value = "";
    domElement.querySelector("#txtStudentName").value = "";
    domElement.querySelector("#txtStudentDept").value = "";
    domElement.querySelector("#txtStudentCity").value = "";
  }

  private async _updateStudent(domElement: any, ID: any) {
    const StudName = domElement.querySelector(`.stud_name${ID}`).textContent;
    const Title = domElement.querySelector(`.stud_id${ID}`).textContent;
    const StudDept = domElement.querySelector(`.stud_dept${ID}`).textContent;
    const StudCity = domElement.querySelector(`.stud_city${ID}`).innerText;

    const lists = pnp.sp.web.lists.getByTitle("Vicky Student List");

    await lists.items
      .getById(ID)
      .update({ Title, StudName, StudDept, StudCity })
      .then((res) => {
        alert("records has been updated!");
      })
      .catch((err) => console.log(err));
  }

  private async _deleteStudent(ID: any) {
    const lists = pnp.sp.web.lists.getByTitle("Vicky Student List");
    await lists.items.getById(ID).delete();
    this._getListData();
  }
  // fetching list by id
  private async _getListById(domElement: any) {
    let studId = domElement.querySelector("#studentId")?.["value"];

    const ListById = await pnp.sp.web.lists
      .getByTitle("Vicky Student List")
      .items.get()
      .then((response: any[]) => {
        return response.find(({ Title, StudName, StudCity, StudDept }: any) => {
          if (Title === studId) {
            return { Title, StudName, StudCity, StudDept };
          }
        });
      });

    if (ListById) {
      this.setState({
        studentList: [ListById],
      });
    } else {
      alert(`oops. ${studId} Don't Exist!`);
    }
  }

  // fetching data crud : get
  private async _getListData() {
    console.log(pnp.sp.web.lists);
    const demo = await pnp.sp.web.lists
      .getByTitle("Vicky Student List")
      .items.get();

    console.log(demo);

    const List = await pnp.sp.web.lists
      .getByTitle("Vicky Student List")
      .items.get()
      .then((response: any[]) => {
        return response.map(
          ({ Title, StudName, StudCity, StudDept, ID }: any) => {
            return { Title, StudName, StudCity, StudDept, ID };
          }
        );
      })
      .catch((err) => console.log(err));

    this.setState({
      studentList: List,
    });
  }
}
