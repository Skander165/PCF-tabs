import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as $ from "jquery";

export class tabs
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private _container: HTMLDivElement;
  private _context: ComponentFramework.Context<IInputs>;
  private static _entityName = "eytn_tabs";
  private _containerTabs: HTMLDivElement;
  private _containerPanels: HTMLDivElement;

  private _getCount: HTMLButtonElement;
  //   private _getCount: EventListener;

  /**
   * Empty constructor.
   */
  constructor() {}

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this._context = context;
    //Create main container
    this._container = document.createElement("div");
    this._container.classList.add("main-container");

    // Create the tabs container
    this._containerTabs = document.createElement("div");
    this._containerTabs.className = "tabs";

    // Create the panels container
    this._containerPanels = document.createElement("div");
    this._containerPanels.className = "panels";

    //Initialize tabs
    this.initTabsAndPanels().then(
      (tabs: ComponentFramework.WebApi.Entity[]) => {
        // The initialization is complete, and tabs contains the array of tabs
        const nbTabs = tabs.length;
        alert("Sucessfully fetched" + nbTabs + "Tabs");
      }
    );

    //Appending containers
    this._container.appendChild(this._containerTabs);
    this._container.appendChild(this._containerPanels);

    //Appending to main container
    container.appendChild(this._container);
  }
  // }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   */
  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this._context = context;
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
   */
  public getOutputs(): IOutputs {
    return {};
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
  }

  private getTabs(): Promise<ComponentFramework.WebApi.Entity[]> {
    // Generate the OData query string to retrieve the tabs
    const queryString = `?$count=true`;

    // Invoke the Web API Retrieve Multiple call
    return this._context.webAPI
      .retrieveMultipleRecords(tabs._entityName, queryString)
      .then((response: ComponentFramework.WebApi.RetrieveMultipleResponse) => {
        // Retrieve Multiple Web API call completed successfully
        const tabs: ComponentFramework.WebApi.Entity[] = response.entities;
        return tabs;
      })
      .catch((errorResponse) => {
        // Error handling code here
        console.log(errorResponse);
        return [];
      });
  }

  private initTabsAndPanels(): Promise<ComponentFramework.WebApi.Entity[]> {
    return this.getTabs().then((tabs: ComponentFramework.WebApi.Entity[]) => {
      const tabElements: HTMLDivElement[] = [];
      const panelElements: HTMLDivElement[] = [];

      for (let i = 0; i < tabs.length; i++) {
        const tab = tabs[i];

        // Create tab element
        const tabElement = document.createElement("div");
        tabElement.className = "tab";
        tabElement.setAttribute("data-tab", `tab-${i}`);
        tabElement.innerText = tab.eytn_name;

        // Create panel element
        const panelElement = document.createElement("div");
        panelElement.className = "panel";
        panelElement.id = `tab-${i}`;
        panelElement.innerText = tab.cr15d_data;

        // Add click event to the tab element
        tabElement.addEventListener("click", () => {
          // Remove the active class from all tabs
          tabElements.forEach((element) => {
            element.classList.remove("active");
          });

          // Add the active class to the clicked tab
          tabElement.classList.add("active");

          // Remove the active class from all panels
          panelElements.forEach((element) => {
            element.classList.remove("active");
          });

          // Add the active class to the corresponding panel
          panelElement.classList.add("active");
        });

        // Append the tab and panel elements to the containers
        this._containerTabs.appendChild(tabElement);
        this._containerPanels.appendChild(panelElement);

        // Store the tab and panel elements in the arrays
        tabElements.push(tabElement);
        panelElements.push(panelElement);
      }

      // Add the active class to the first tab and panel initially
      if (tabElements.length > 0 && panelElements.length > 0) {
        tabElements[0].classList.add("active");
        panelElements[0].classList.add("active");
      }

      return tabs;
    });
  }
}
