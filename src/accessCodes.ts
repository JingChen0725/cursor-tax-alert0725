import type { AccessRole } from "./accessSession";

export type AccessCodeProfile = {
  role: AccessRole;
  salesperson: string;
};

export const ACCESS_CODE_MAP: Record<string, AccessCodeProfile> = {
  QLNXVBTMRAKCP: { role: "admin", salesperson: "管理员" },
  VJZQTPKLMRSHD: { role: "sales", salesperson: "陈炳峰" },
  MNXKQTRVPLSAD: { role: "sales", salesperson: "陈国忠" },
  TQVRMSKLPXJAD: { role: "sales", salesperson: "陈金雍" },
  PKLMSQTRVNXAD: { role: "sales", salesperson: "陈文祥" },
  ZTRQKLPMSVNAD: { role: "sales", salesperson: "黄汉民" },
  RVKQPLMTSNXAD: { role: "sales", salesperson: "黄苏航" },
  SKQTRVPLMNXAD: { role: "sales", salesperson: "王国秋" },
  NQTRVPLKMSXAD: { role: "sales", salesperson: "吴学晃" },
  XQTRVPLKMSNAD: { role: "sales", salesperson: "卓炜琪" },
};
