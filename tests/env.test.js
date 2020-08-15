const { expect } = require("chai");

describe("env check", () => {
  it("hasBuffer", () => {
    const hasBuffer =
      typeof Buffer !== "undefined" &&
      typeof process !== "undefined" &&
      typeof process.versions !== "undefined" &&
      !!process.versions.node;

    expect(hasBuffer).to.equal(true);
  });

  it("typeof exports !== 'undefined'", () => {
    expect(typeof exports !== "undefined").to.equal(true);
  });

  it("typeof module !== 'undefined' && module.exports'", () => {
    expect(typeof module !== "undefined" && !!module.exports).to.equal(true);
  });
});
