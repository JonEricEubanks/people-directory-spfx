  import * as React from 'react';
  import { useEffect, useMemo, useRef, useState } from 'react';
  import { ThemeProvider, IReadonlyTheme } from '@microsoft/sp-component-base';
  import { WebPartTitle } from '@pnp/spfx-controls-react';
  import {
    Dropdown,
    Option,
    OptionOnSelectData,
    Overflow,
    OverflowItem,
    SearchBox,
    Field,
    Title2,
    MessageBar,
    MessageBarBody,
    MessageBarTitle,
    Tab,
    TabList,
    tokens,
    makeStyles,
    mergeClasses,
    SelectTabEvent,
    SelectTabData,
  } from '@fluentui/react-components';
  import { People48Filled } from '@fluentui/react-icons';

  import { ISPServices } from '../../../SPServices/ISPServices';
  import { IDirectoryProps } from './IDirectoryProps';
  import Paging from './Pagination/Paging';
  import { OverflowAlphabetsMenu } from './OverflowAlphabetsMenu/OverflowAlphabetsMenu';
  import styles from './Directory.module.scss';
  import { PersonaCard } from './PersonaCard/PersonaCard';
  import { spservices } from '../../../SPServices/spservices';
  import { IDirectoryState } from './IDirectoryState';
  import * as strings from 'DirectoryWebPartStrings';
  import { Shimmer } from './Shimmer/Shimmer';
  import { IStackTokens, Stack } from '@fluentui/react';

  const wrapStackTokens: IStackTokens = { childrenGap: 30 };

  const useFluentStyles = makeStyles({
    alphabets: {
      overflow: 'hidden',
      padding: '4px',
      paddingTop: '10px',
      zIndex: 0,
      borderRadius: '0',
      borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    },
    horizontal: {
      height: 'fit-content',
      minWidth: '100%',
    },
    tabList: {
      justifyContent: 'center',
    },
  });

  const DirectoryHook: React.FC<IDirectoryProps> = (props) => {
    const _services: ISPServices = new spservices(props.context);
    const [az, setaz] = useState<string[]>([]);
    const [alphaKey, setalphaKey] = useState<string>('A');
    const [state, setstate] = useState<IDirectoryState>({
      users: [],
      isLoading: true,
      errorMessage: '',
      hasError: false,
      indexSelectedKey: 'A',
      searchString: 'LastName',
      searchText: '',
    });
    const orderOptions = [
      { value: 'FirstName', text: 'First Name' },
      { value: 'LastName', text: 'Last Name' },
      { value: 'Department', text: 'Department' },
      { value: 'Location', text: 'Location' },
      { value: 'JobTitle', text: 'Job Title' },
    ];
    const color = props.context.microsoftTeams ? 'white' : '';
    const searchVersionRef = useRef(0);

    // Paging
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const [pagedItems, setPagedItems] = useState<any[]>([]);
    const [pageSize, setPageSize] = useState<number>(
      props.pageSize ? props.pageSize : 10
    );
    const [currentPage, setCurrentPage] = useState<number>(1);

    const _onPageUpdate = async (pageno?: number): Promise<void> => {
      const currentPge = pageno ? pageno : currentPage;
      const startItem = (currentPge - 1) * pageSize;
      const endItem = currentPge * pageSize;
      const filItems = state.users.slice(startItem, endItem);
      setCurrentPage(currentPge);
      setPagedItems(filItems);
    };

    const diretoryGrid =
      pagedItems && pagedItems.length > 0
        ? // eslint-disable-next-line @typescript-eslint/no-explicit-any
          pagedItems.map((user: any, i) => {
            return (
              <PersonaCard
                context={props.context}
                key={'PersonaCard' + i}
                cardIndex={i}
                profileProperties={{
                  DisplayName: user.PreferredName,
                  Title: user.JobTitle,
                  PictureUrl: user.PictureURL,
                  Email: user.WorkEmail,
                  Department: user.Department,
                  WorkPhone: user.WorkPhone,
                  Location: user.OfficeNumber
                    ? user.OfficeNumber
                    : user.BaseOfficeLocation,
                }}
              />
            );
          })
        : [];
    const _loadAlphabets = (): void => {
      const alphabets: string[] = [];
      for (let i = 65; i < 91; i++) {
        alphabets.push(String.fromCharCode(i));
      }
      setaz(alphabets);
    };

    const _alphabetChange = async (event: SelectTabEvent|string, data: SelectTabData | undefined): Promise<void> => {
      //check if event is string or SelectTabEvent
      let value: string = '';
      if (typeof event === 'string') {
        value = event;
      } else {
        value = event?.currentTarget.innerText || '';
      }
      searchVersionRef.current++;
      setstate((prev) => ({
        ...prev,
        searchText: '',
        indexSelectedKey: value,
        isLoading: true,
      }));
      setalphaKey(value);
      setCurrentPage(1);
    };
    const _searchByAlphabets = async (initialSearch: boolean): Promise<void> => {
      const myVersion = searchVersionRef.current;
      setstate((prev) => ({ ...prev, isLoading: true }));
      let users = null;
      if (initialSearch) {
        if (props.searchFirstName)
          users = await _services.searchUsersNew('', `FirstName:a*`, false);
        else users = await _services.searchUsersNew('a', '', true);
      } else {
        if (props.searchFirstName)
          users = await _services.searchUsersNew(
            '',
            `FirstName:${alphaKey}*`,
            false
          );
        else users = await _services.searchUsersNew(`${alphaKey}`, '', true);
      }
      // Discard if user has typed or clicked a letter since this call started
      if (myVersion !== searchVersionRef.current) return;
      setstate((prev) => ({
        ...prev,
        // Never overwrite searchText from async completion — the input owns its own value
        indexSelectedKey: initialSearch ? 'A' : prev.indexSelectedKey,
        users:
          users && users.PrimarySearchResults ? users.PrimarySearchResults : null,
        isLoading: false,
        errorMessage: '',
        hasError: false,
      }));
    };

    const _searchUsers = async (searchText: string): Promise<void> => {
      // Snapshot the version that was set when the user typed; don't increment
      // here because _searchBoxChanged already did it on each keystroke.
      const myVersion = searchVersionRef.current;
      try {
        setstate((prev) => ({ ...prev, isLoading: true }));
        if (searchText.length > 0) {
          const searchProps: string[] =
            props.searchProps && props.searchProps.length > 0
              ? props.searchProps.split(',')
              : ['FirstName', 'LastName', 'Department'];
          let qryText = '';
          const finalSearchText: string = searchText
            ? searchText.replace(/ /g, '+')
            : searchText;
          if (props.clearTextSearchProps) {
            const tmpCTProps: string[] =
              props.clearTextSearchProps.indexOf(',') >= 0
                ? props.clearTextSearchProps.split(',')
                : [props.clearTextSearchProps];
            if (tmpCTProps.length > 0) {
              searchProps.map((srchprop, index) => {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const ctPresent: any[] = tmpCTProps.filter((o) => {
                  return o.toLowerCase() === srchprop.toLowerCase();
                });
                if (ctPresent.length > 0) {
                  if (index === searchProps.length - 1) {
                    qryText += `${srchprop}:${searchText}*`;
                  } else qryText += `${srchprop}:${searchText}* OR `;
                } else {
                  if (index === searchProps.length - 1) {
                    qryText += `${srchprop}:${finalSearchText}*`;
                  } else qryText += `${srchprop}:${finalSearchText}* OR `;
                }
              });
            } else {
              searchProps.map((srchprop, index) => {
                if (index === searchProps.length - 1)
                  qryText += `${srchprop}:${finalSearchText}*`;
                else qryText += `${srchprop}:${finalSearchText}* OR `;
              });
            }
          } else {
            searchProps.map((srchprop, index) => {
              if (index === searchProps.length - 1)
                qryText += `${srchprop}:${finalSearchText}*`;
              else qryText += `${srchprop}:${finalSearchText}* OR `;
            });
          }
          console.log(qryText);
          const users = await _services.searchUsersNew('', qryText, false);
          if (myVersion !== searchVersionRef.current) return;
          setstate((prev) => ({
            ...prev,
            indexSelectedKey: '0',
            users:
              users && users.PrimarySearchResults
                ? users.PrimarySearchResults
                : null,
            isLoading: false,
            errorMessage: '',
            hasError: false,
          }));
          setalphaKey('0');
        } else {
          await _searchByAlphabets(true);
        }
      } catch (err) {
        if (myVersion !== searchVersionRef.current) return;
        const errorMessage = err instanceof Error ? err.message : String(err);
        setstate((prev) => ({ ...prev, errorMessage, hasError: true }));
      }
    };

    const _searchBoxChanged = (newvalue: string): void => {
      // Increment version immediately so any in-flight search is invalidated
      // before the 300ms debounce even fires.
      searchVersionRef.current++;
      setCurrentPage(1);
      setstate((prev) => ({ ...prev, searchText: newvalue }));
    };

    useEffect(() => {
      const debouncedSearch = setTimeout(() => {
        if (state.searchText !== undefined) {
          _searchUsers(state.searchText).catch((err) => {});
        }
      }, 300);

      return () => clearTimeout(debouncedSearch);
    }, [state.searchText]);

    const _sortPeople = async (sortField: string): Promise<void> => {
      let _users = [...state.users];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      _users = _users.sort((a: any, b: any) => {
        switch (sortField) {
          // Sort by Location
          case 'Location':
            if (
              (a.BaseOfficeLocation || '').toUpperCase() <
              (b.BaseOfficeLocation || '').toUpperCase()
            ) {
              return -1;
            }
            if (
              (a.BaseOfficeLocation || '').toUpperCase() >
              (b.BaseOfficeLocation || '').toUpperCase()
            ) {
              return 1;
            }
            return 0;

          default:
            if (
              (a[sortField] || '').toUpperCase() <
              (b[sortField] || '').toUpperCase()
            ) {
              return -1;
            }
            if (
              (a[sortField] || '').toUpperCase() >
              (b[sortField] || '').toUpperCase()
            ) {
              return 1;
            }
            return 0;
        }
      });
      setstate({ ...state, users: _users, searchString: sortField });
    };

    useEffect(() => {
      setPageSize(props.pageSize);
      if (state.users) {
        _onPageUpdate().catch((err) => {});
      }
    }, [state.users, props.pageSize]);

    useEffect(() => {
      if (alphaKey.length > 0 && alphaKey !== '0') _searchByAlphabets(false).catch((err) => {});
    }, [alphaKey]);

    useEffect(() => {
      _loadAlphabets();
      _searchByAlphabets(true).catch((err) => {});
    }, [props]);

    const onOptionSelect = (ev: React.FormEvent<HTMLElement>, data: OptionOnSelectData): void => {
      _sortPeople(data.optionValue).catch((err) => {});
    };

    // Add the new callback function
    const handleSearchKeyPress = React.useCallback(
      (ev: React.KeyboardEvent<HTMLInputElement>) => {
        if (ev.key === 'Enter') {
          _searchUsers(state.searchText).catch((err) => {});
        }
      },
      [state.searchText]
    );

    const fluentStyles = useFluentStyles();

    const themeVars = useMemo(() => {
      const tp = props.context.serviceScope.consume(ThemeProvider.serviceKey);
      const theme: IReadonlyTheme | undefined = tp.tryGetTheme();
      const p = theme?.palette;
      return {
        '--spd-theme-primary': p?.themePrimary || '#0078d4',
        '--spd-theme-dark': p?.themeDark || '#005a9e',
        '--spd-theme-darker': p?.themeDarker || '#004578',
        '--spd-theme-light': p?.themeLight || '#c7e0f4',
      } as React.CSSProperties;
    }, [props.context]);

    return (
      <div className={styles.directory} style={themeVars}>
        <WebPartTitle
          displayMode={props.displayMode}
          title={props.title}
          updateProperty={props.updateProperty}
        />
        <div className={styles.searchBox}>
          <SearchBox
            type="search"
            placeholder={strings.SearchPlaceHolder}
            className={styles.searchTextBox}
            value={state.searchText}
            onKeyDown={handleSearchKeyPress}
            onChange={(_, data) => _searchBoxChanged(data.value)}
          />
          <div
            className={mergeClasses(
              fluentStyles.alphabets,
              fluentStyles.horizontal
            )}
          >
            <Overflow minimumVisible={2}>
              <TabList
                selectedValue={state.indexSelectedKey}
                // appearance="filled-circular"
                onTabSelect={_alphabetChange}
                className={fluentStyles.tabList}
              >
                {az.map((index: string) => {
                  return (
                    <OverflowItem key={index} id={index}>
                      <Tab value={index} key={index}>
                        {index}
                      </Tab>
                    </OverflowItem>
                  );
                })}
                <OverflowAlphabetsMenu onTabSelect={(value)=>_alphabetChange(value,undefined)} tabs={az} />
              </TabList>
            </Overflow>
          </div>
        </div>
        {state.isLoading ? (
          <div style={{ marginTop: '10px' }}>
            <Shimmer />
          </div>
        ) : (
          <>
            {state.hasError ? (
              <div style={{ marginTop: '10px' }}>
                <MessageBar intent="error">
                  <MessageBarBody>
                    <MessageBarTitle>{state.errorMessage}</MessageBarTitle>
                  </MessageBarBody>
                </MessageBar>
              </div>
            ) : (
              <>
                {!pagedItems || pagedItems.length === 0 ? (
                  <div className={styles.noUsers}>
                    <People48Filled style={{ fontSize: '54px', color: color }} />
                    <Title2
                      style={{
                        marginLeft: 5,
                        color: color,
                      }}
                    >
                      {strings.DirectoryMessage}
                    </Title2>
                  </div>
                ) : (
                  <>
                    <div style={{ width: '100%', display: 'inline-block' }}>
                      <Paging
                        totalItems={state.users.length}
                        itemsCountPerPage={pageSize}
                        onPageUpdate={_onPageUpdate}
                        currentPage={currentPage}
                      />
                    </div>
                    <div className={styles.dropDownSortBy}>
                      <Stack
                        horizontal
                        horizontalAlign="center"
                        wrap
                        tokens={wrapStackTokens}
                      >
                        <Field label={strings.DropDownPlaceLabelMessage}>
                          <Dropdown
                            placeholder={strings.DropDownPlaceHolderMessage}
                            value={state.searchString}
                            onOptionSelect={onOptionSelect}
                          >
                            {orderOptions.map((option: { value: string; text: string }) => {
                              return (
                                <Option key={option.value} value={option.value}>
                                  {option.text}
                                </Option>
                              );
                            })}
                          </Dropdown>
                        </Field>
                      </Stack>
                    </div>
                    <div className={styles.cardsGrid}>
                      {diretoryGrid}
                    </div>
                    <div style={{ width: '100%', display: 'inline-block' }}>
                      <Paging
                        totalItems={state.users.length}
                        itemsCountPerPage={pageSize}
                        onPageUpdate={_onPageUpdate}
                        currentPage={currentPage}
                      />
                    </div>
                  </>
                )}
              </>
            )}
          </>
        )}
      </div>
    );
  };

  export default DirectoryHook;
