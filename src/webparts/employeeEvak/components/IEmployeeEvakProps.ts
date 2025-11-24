export interface IEmployeeEvakProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
import { SPFI } from "@pnp/sp";

export interface IEmployeeEvakProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  // ⭐ NEW
  sp: SPFI;
}
