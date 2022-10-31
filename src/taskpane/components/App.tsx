import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { ReportDataString } from "../data";
import * as Papa from "papaparse";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
  writeDataCompleteMessage: string;
}

export default class App extends React.Component<AppProps, AppState> {
  parseResult: Papa.ParseResult<any>;
  data: string[][];

  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      writeDataCompleteMessage: "",
    };

    this.parseResult = Papa.parse(ReportDataString);
    this.data = this.parseResult.data;
  }

  click = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  createTable = async () => {
    try {
      await Excel.run({ delayForCellEdit: true }, async (context: Excel.RequestContext) => {
        context.application.suspendScreenUpdatingUntilNextSync();
        let worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.getUsedRange().clear();
        const headerRange: Excel.Range = worksheet.getRange("A1:AJ1");

        headerRange.values = [
          [
            "ID",
            "Source",
            "Number",
            "cvbvbvbvbvbv",
            "nvbcxvbcvbx",
            "Product",
            "Location",
            "Quantity",
            "Value",
            "Person",
            "Guy",
            "Start Date",
            "End Date",
            "Number",
            "Transaction ID",
            "S",
            "Date",
            "External ID",
            "User",
            "Trans From Date",
            "Trans To Date",
            "werwerwrwer",
            "gsdfgdsfg",
            "Period",
            "Period Start",
            "Period End",
            "Child Product",
            "Currency",
            "yyyyy",
            "iiiiii",
            "time",
            "wwww",
            "eeeee",
            "ppppp",
            "bbbbb",
            "hhhhh",
          ],
        ];

        worksheet.tables.add(headerRange, true);
        return context.sync();
      }).catch((e) => {
        console.log(`Error creating report table in Excel:`);
        console.log(e);
      });
    } catch (e) {
      console.log(`Error creating report table:`);
      console.log(e);
    }
  };

  writeReportData = (totalColumnsEndString: string) => {
    let rowCount = this.data.length;
    Excel.run({ delayForCellEdit: true }, async (context: Excel.RequestContext): Promise<any> => {
      const tableBodyLastRowIndex: number = 1 + rowCount;
      let reportTable: Excel.Table = context.workbook.worksheets.getActiveWorksheet().tables.getItemAt(0);
      reportTable.resize(reportTable.worksheet.getRange(`A1:${totalColumnsEndString}${tableBodyLastRowIndex}`));

      const bite: number = 1000;
      for (let i = 0; i < rowCount; i += bite) {
        const rowData: string[][] = this.data.splice(0, bite);
        const tableRow: Excel.Range = reportTable.worksheet.getRange(
          `A${i + 2}:${totalColumnsEndString}${i + rowData.length + 1}`
        );
        tableRow.values = rowData;
        tableRow.untrack();

        context.sync();
      }

      // Format table
      reportTable.getHeaderRowRange().format.autofitColumns();
      return context.sync().then(() => this.setState({ writeDataCompleteMessage: "Completed writing report data." }));
    });
  };

  populateTable = async () => {
    return Excel.run({ delayForCellEdit: true }, async (context: Excel.RequestContext): Promise<any> => {
      const reportTable = context.workbook.worksheets.getActiveWorksheet().tables.getItemAt(0);
      reportTable.load("isNullObject, id");
      await context.sync();
      if (reportTable.isNullObject) {
        return context.sync();
      }

      reportTable.getDataBodyRange().untrack().numberFormat = [
        [
          "@",
          "@",
          "@",
          "@",
          "@",
          "@",
          "@",
          "#,##0",
          "#,##0.0000",
          "@",
          "@",
          "m/d/yyyy h:mm AM/PM",
          "m/d/yyyy h:mm AM/PM",
          "@",
          "@",
          "@",
          "m/d/yyyy",
          "@",
          "@",
          "m/d/yyyy",
          "m/d/yyyy",
          "@",
          "@",
          "@",
          "@",
          "@",
          "@",
          "@",
          "#,##0.0000",
          "#,##0.0000",
          "@",
          "@",
          "@",
          "@",
          "@",
          "@",
        ],
      ];

      this.writeReportData("AJ");

      return context.sync().catch((error) => console.log(error));
    }).catch((error) => console.log(error));
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
        <HeroList
          message="First press Create Table to generate the table. Then press Populate Table to write the data to the table."
          items={this.state.listItems}
        >
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.createTable}
          >
            Create Table
          </DefaultButton>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.populateTable}
          >
            Populate Table
          </DefaultButton>
          <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">
            {this.state.writeDataCompleteMessage}
          </h2>
        </HeroList>
      </div>
    );
  }
}
