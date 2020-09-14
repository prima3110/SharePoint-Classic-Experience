class DepartmentsWP {
  constructor() {
    this.departmentsList;
    this.departmentsBodyId = "dprtBody";
    this.listTitle = "MyContacts";
    this.endPoint = `_api/web/lists/getbytitle('${this.listTitle}')/items`;
    this.url = `${_spPageContextInfo.webAbsoluteUrl}/`; // url сайт-колекції
    this.queryAllItems = `${this.url}${this.endPoint}`; // строка запиту для отримання всіх items по title
  }

  async getInitialItems() {
    // отримуємо початковий список айтемів і рендеримо їх на сторінку
    try {
      const items = await this.getItems(this.queryAllItems);
      const btn = document.querySelector(".btn");
      btn.addEventListener("click", () => this.postNewItem());
      this.renderHTML(items); // рендеримо сторінку
      return items;
    } catch (error) {
      console.log("error in getData", error);
    }
  }

  async createNewItem() {
    // створюємо новий айтем
    try {
      const requestDigest = await this.getRequestDigest(this.url); // отримали X-RequestDigest
      const listItemType = await this.getListItemType(this.url, this.listTitle); // отримали listItemType поточного списку

      const newItem = {
        Title: "newTitle",
        sbDepartmentIconURL: {
          Description:
            "https://opendatasecurity.io/wp-content/uploads/2017/05/can-you-hire-a-hacker-ods.jpg",
          Url:
            "https://opendatasecurity.io/wp-content/uploads/2017/05/can-you-hire-a-hacker-ods.jpg",
        },
        sbDepartmentURL: {
          Description:
            "https://opendatasecurity.io/wp-content/uploads/2017/05/can-you-hire-a-hacker-ods.jpg",
          Url:
            "https://opendatasecurity.io/wp-content/uploads/2017/05/can-you-hire-a-hacker-ods.jpg",
        },
      }; // новий об'єкт, який будемо додавати

      const objType = {
        __metadata: {
          type: listItemType.d.ListItemEntityTypeFullName,
        },
      }; // вказуємо тип нового об'єкту

      const objData = JSON.stringify(Object.assign(objType, newItem)); // оновлений об'єкт з типом

      return $.ajax({
        url: this.queryAllItems,
        type: "POST",
        data: objData,
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest":
            requestDigest.d.GetContextWebInformation.FormDigestValue,
          "X-HTTP-Method": "POST",
        },
      });
    } catch (error) {
      console.log("error in createNewItem", error);
    }
  }

  getItems(query) {
    // отримуємо айтеми по title
    try {
      const items = $.ajax({
        url: query,
        method: "GET",
        contentType: "application/json;odata=verbose",
        headers: {
          Accept: "application/json;odata=verbose",
        },
      });
      return items;
    } catch (error) {
      console.log("error in getItems", error);
    }
  }

  getRequestDigest(webUrl) {
    // отримуємо X-RequestDigest
    try {
      const requestDigest = $.ajax({
        url: webUrl + "_api/contextinfo",
        method: "POST",
        headers: {
          Accept: "application/json; odata=verbose",
        },
      });
      return requestDigest;
    } catch (error) {
      console.log("error in getRequestDigest", error);
    }
  }

  getListItemType(url, listTitle) {
    // отримуємо listItemType поточного списку
    try {
      const queryListItemType =
        url +
        "_api/Web/Lists/getbytitle('" +
        listTitle +
        "')/ListItemEntityTypeFullName";
      return this.getItems(queryListItemType);
    } catch (error) {
      console.log("error in getListItemType", error);
    }
  }

  async postNewItem() {
    // додаємо новий айтем
    try {
      const newItem = await this.createNewItem();
      console.log("newItem", newItem);
    } catch (error) {
      console.log("error", error);
    }
  }

  renderHTML(result) {
    try {
      const departmentsList = result.d.results;
      let departmentsItems = "";
      departmentsList.map((item) => {
        const departmentItem = new DepartmentItem(item);
        departmentsItems += departmentItem.getHTML();
      });
      document.getElementById(
        this.departmentsBodyId
      ).innerHTML = departmentsItems;

      $("#dprtBody").click(function (event) {
        const $elem = $(event.target);
        if ($elem.closest(".dprtCard")[0]) {
          window.open($elem.closest(".dprtCard").data("url"));
        }
      });
    } catch (error) {
      this.webpartNoDataContainer("dprtWrapper", error);
    }
  }

  webpartNoDataContainer(domName, errorMessage) {
    let errorHTML = `<div class="noDataContainer">
        <div id="warning-block">
            <img src="../SiteAssets/departments-sandbox-ropry/noData.svg" alt="error"/>
        </div>
      </div>`;
    document.getElementById(domName).innerHTML = errorHTML;
    console.error(`Error in ${domName} -> ${errorMessage}`);
  }
}

class DepartmentItem {
  constructor(departmentItem) {
    this.img = departmentItem.sbDepartmentIconURL
      ? departmentItem.sbDepartmentIconURL.Url
      : "";
    this.title = departmentItem.Title;
    this.web = departmentItem.sbDepartmentURL
      ? departmentItem.sbDepartmentURL.Url
      : "";
    this.phone = departmentItem.CellPhone || "hidden";
  }

  getHTML() {
    let departmentItemTemplate = `
      <div class="dprtCard dprtCard_hover" data-url=${this.web}>
        <img class="dprtCard__img" src=${this.img}  alt="photo"/>
        <h3 class="dprtCard__name">${this.title}</h3>
        <span class="dprtCard__web dprtCard__web_visited"><img class="dprtCard__icon" src="${_spPageContextInfo.siteServerRelativeUrl}/SiteAssets/departments-sandbox-ropry/phone.webp" alt="icon" width=30>&nbsp; ${this.phone} </span>
      </div>`;
    return departmentItemTemplate;
  }
}

SP.SOD.executeFunc("sp.js", "SP.ClientContext", function () {
  const dprtWP = new DepartmentsWP();
  dprtWP.getInitialItems();
});
