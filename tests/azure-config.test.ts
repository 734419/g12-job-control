import { describe, it, expect } from "vitest";

describe("Azure AD Configuration", () => {
  it("should have EXPO_PUBLIC_AZURE_CLIENT_ID set to a valid GUID", () => {
    const clientId = process.env.EXPO_PUBLIC_AZURE_CLIENT_ID;
    expect(clientId).toBeDefined();
    expect(clientId).toMatch(
      /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
    );
  });

  it("should have EXPO_PUBLIC_AZURE_TENANT_ID set to a valid GUID", () => {
    const tenantId = process.env.EXPO_PUBLIC_AZURE_TENANT_ID;
    expect(tenantId).toBeDefined();
    expect(tenantId).toMatch(
      /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
    );
  });

  it("should be able to construct a valid MSAL authority URL", () => {
    const tenantId = process.env.EXPO_PUBLIC_AZURE_TENANT_ID;
    const authority = `https://login.microsoftonline.com/${tenantId}`;
    expect(authority).toBe(
      "https://login.microsoftonline.com/58fe4fb8-9154-4e1f-acb1-b568e951d7b3"
    );
  });
});
