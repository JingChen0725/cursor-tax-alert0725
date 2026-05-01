import { createContext, useContext } from "react";

export type AccessRole = "admin" | "sales";

export type AccessSession = {
  code: string;
  role: AccessRole;
  salesperson: string;
};

export const AccessSessionContext = createContext<AccessSession | null>(null);

export function useAccessSession(): AccessSession | null {
  return useContext(AccessSessionContext);
}
