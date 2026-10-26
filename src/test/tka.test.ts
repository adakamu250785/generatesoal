import { describe, it, expect } from "vitest";
import { isNumeracySubject, tkaInstruction } from "@/lib/tka";

describe("standar TKA", () => {
  it("Pendidikan Pancasila tidak memakai numerasi", () => {
    expect(isNumeracySubject("Pendidikan Pancasila")).toBe(false);
    expect(tkaInstruction("Pendidikan Pancasila")).toContain("JANGAN memaksakan angka");
  });
  it("Matematika memakai literasi dan numerasi", () => {
    expect(isNumeracySubject("Matematika")).toBe(true);
    expect(tkaInstruction("Matematika")).not.toContain("JANGAN memaksakan angka");
  });
});
