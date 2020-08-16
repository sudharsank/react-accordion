import { DisplayMode } from '@microsoft/sp-core-library';

export interface IReactAccordionProps {
  listId: string;
  title: string;
  displayMode: DisplayMode;
  maxItemsPerPage: number;
  updateProperty: (value: string) => void;
  configurePropertyPane: () => void;
}
