import * as React from "react";

import {
  Combobox,
  Label,
  OptionGroup,
  Option,
  makeStyles,
  SelectionEvents,
  OptionOnSelectData,
  FluentProvider,
  Theme,
  Input,
} from "@fluentui/react-components";
import { ILookupToCChoiceProps, IRecord, IRecordCategory, IMru } from "../Interfaces/Globals";
import {
  StarRegular,
  ClockRegular,
  AddRegular,
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  root: {
    width: "100%"
  },
});

export interface IILookupToChoiceState{
  categories : IRecordCategory[],
  entityIdFieldName: string,
  entityNameFieldName: string,
  entityDisplayName: string,
  selectedKey: string,
  selectedText: string,
  parentRecordId : string | undefined,
}

export const LookupCombobox = (props: ILookupToCChoiceProps): JSX.Element => {
  const [state, setState] = React.useState({
    categories: [] as IRecordCategory[],
    entityIdFieldName: "",
    entityNameFieldName: "",
    entityDisplayName: "",
    selectedKey: props.selectedId,
    selectedText: "",
    parentRecordId : props.parentRecordId,
  } as IILookupToChoiceState);

  const [query, setQuery] = React.useState<string>("");
  const [isSearching, setIsSearching] = React.useState<boolean>(false);

  let availableOptions : IRecord[] = [];

  const styles = useStyles();
  const currentTheme = props.context.fluentDesignLanguage?.tokenTheme as Theme;
  const myTheme = props.isDisabled
    ? {
        ...currentTheme,
        colorCompoundBrandStroke: currentTheme?.colorNeutralStroke1,
        colorCompoundBrandStrokeHover: currentTheme?.colorNeutralStroke1Hover,
        colorCompoundBrandStrokePressed:
          currentTheme?.colorNeutralStroke1Pressed,
        colorCompoundBrandStrokeSelected:
          currentTheme?.colorNeutralStroke1Selected,
      }
    : currentTheme;

  React.useEffect(() => {
    retrieveMetadata();
  },[]);
 
  React.useEffect(() => {
    retrieveRecords();
  }, [state.entityIdFieldName]);

  React.useEffect(() => {
    retrieveRecords();
  }, [props.parentRecordId]);

  React.useEffect(() => {
    displayRecords(state.categories, props.selectedId);
 }, [props.selectedId]);

 React.useEffect(() => {
  displayRecords(state.categories, state.selectedKey);
}, [state.selectedKey]);

  const retrieveMetadata = () => {
    props.context.utils
      .getEntityMetadata(props.entityName)
      .then((metadata) => {
        setState((prevState) => {
          return {
            ...prevState,
            entityIdFieldName: metadata.PrimaryIdAttribute as string,
            entityNameFieldName: metadata.PrimaryNameAttribute as string,
            entityDisplayName: metadata.DisplayName as string,
          };
        });
        return null;
      })
      .catch((error) => {
        console.log(error);
      });
  };


  const setRecords = (categories: IRecordCategory[], records: IRecord[]) => {
    const category = categories.find((c) => c.key === "records");
    if (!category) {
      categories.push({
        title: state.entityDisplayName,
        key: "records",
        type: "records",
        records: records,
      });
    } else {
      category.title = state.entityDisplayName;
      category.records = records;
    }
  };

  const sortRecords = (records: IRecord[]) => {
    records = records.sort((n1, n2) => {
      if (n1.text.toLowerCase() > n2.text.toLowerCase()) {
        return 1;
      }

      if (n1.text.toLowerCase() < n2.text.toLowerCase()) {
        return -1;
      }

      return 0;
    });
  };


  const retrieveRecords = () => {

    if(!(state.entityDisplayName?.length > 0)){
      return;
    }

    let filter = "";
    if (props.viewId) {
      filter =
        "?$top=1&$select=fetchxml,returnedtypecode&$filter=savedqueryid eq " +
        props.viewId;
    } else {
      filter =
        "?$top=1&$select=fetchxml,returnedtypecode&$filter=returnedtypecode eq '" +
        props.entityName +
        "' and querytype eq 64";
    }

    props.context.webAPI
      .retrieveMultipleRecords("savedquery", filter)
      .then((result) => {
        const view = result.entities[0];
        const xml = view.fetchxml as string;

        props.context.webAPI
          .retrieveMultipleRecords(
            view.returnedtypecode as string,
            "?fetchXml=" + xml
          )
          .then((result) => {
            availableOptions = result.entities.map((r) => {
              let localizedEntityFieldName = "";
              const mask = props.context.parameters.attributemask.raw;

              if (mask) {
                localizedEntityFieldName = mask.replace(
                  "{lcid}",
                  props.context.userSettings.languageId.toString()
                );
              }

              return {
                key: r[state.entityIdFieldName] as string,
                text: (r[localizedEntityFieldName] ??
                  r[state.entityNameFieldName] ??
                  "Display Name is not available") as string,
              };
            });

            if (props.context.parameters.sortByName.raw === "1") {
              sortRecords(availableOptions);
            }

            availableOptions.splice(0, 0, { key: "---", text: "---" });

            const categories = [] as IRecordCategory[];
            setRecords(categories, availableOptions);
            displayRecords(categories, props.selectedId);

            return null;
          })
          .catch((error) => {
            console.log((error as Error).message);
          });

        return null;
      })
      .catch((error) => {
        console.log((error as Error).message);
      });
  };

  const displayRecords = (categories : IRecordCategory[], selectedKey: string) => {

    availableOptions = availableOptions.length > 0 ? availableOptions : state.categories.find((c) => c.type === "records")?.records ?? []

    if(availableOptions.length === 0){
      return;}

    const selectedOption = availableOptions.find(
      (o) => o.key === selectedKey
    );

    setQuery(selectedOption?.text ?? "---");
    setState((prevState) => {
      return {
        ...prevState,
        categories: categories,
        selectedKey: selectedOption?.key ?? "---",
        selectedText: selectedOption?.text ?? "---",
      };
    });
  }

  const updateSelectedItem = (
    selectedId: string | undefined,
    selectedText: string | undefined
  ) => {
    setState((prevState) => ({
      ...prevState,
      selectedKey: selectedId?.replace(/[{}]/g,"").toLowerCase() ?? "---",
      selectedText: selectedText ?? "---",
    }));
    setIsSearching(false);
    setQuery(selectedText ?? "");
  };

  const handleOptionSelect = (
    event: SelectionEvents,
    data: OptionOnSelectData
  ) => {
      const newValue = {
        id: (data.optionValue ?? "").split("_mru")[0].split("_fav")[0],
        name: data.optionText,
        entityType: props.entityName,
      };
      props.notifyOutputChanged(newValue);
      updateSelectedItem(newValue.id, newValue.name);
  };

  return (
    <div className={styles.root}>
      <FluentProvider theme={myTheme} className={styles.root}>
        {props.isDisabled ? (
          <Input
            value={state.selectedText}
            appearance="filled-darker"
            className={styles.root}
            readOnly={props.isDisabled}
          />
        ) : (
          <Combobox
            {...props}
            placeholder="---"
            onChange={(ev) => {
              setQuery(ev.target.value);
              setIsSearching(true);
            }}
            onOptionSelect={handleOptionSelect}
            selectedOptions={[state.selectedKey]}
            value={query}
            className={styles.root}
            appearance="filled-darker"
          >
            {state.categories.length === 1
              ? state.categories[0].records
                  .filter(
                    (r) =>
                      (isSearching &&
                        r.text.toLowerCase().includes(query.toLowerCase())) ||
                      !isSearching
                  )
                  .map((record) => (
                    <Option
                      key={record.key}
                      text={record.text}
                      value={record.key}
                      className="{styles.root}"
                    >
                      {record.text}
                    </Option>
                  ))
              : state.categories.map((category) => (
                  <OptionGroup
                    label={category.title}
                    key={category.key}
                    className="{styles.root}"
                  >
                    <div
                      style={
                        category.key === "records"
                          ? {
                              maxHeight: "300px",
                              overflowY: "auto",
                              overflowX: "hidden",
                            }
                          : {}
                      }
                    >
                      {category.records
                        .filter(
                          (r) =>
                            (isSearching &&
                              r.text
                                .toLowerCase()
                                .includes(query.toLowerCase()) &&
                              category.type === "records") ||
                            !isSearching ||
                            category.type != "records"
                        )
                        .map((record) => (
                          <Option
                            key={record.key}
                            text={record.text}
                            value={record.key}
                            className="{styles.root}"
                          >
                            <Label>{record.text}</Label>
                          </Option>
                        ))}
                    </div>
                  </OptionGroup>
                ))}
          </Combobox>
        )}
      </FluentProvider>
    </div>
  );
};
