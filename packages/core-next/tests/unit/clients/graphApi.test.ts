/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, beforeEach, afterEach } from "mocha";
import * as sinon from "sinon";
import axios from "axios";
import * as retryModule from "../../../src/http/retry";
import { createMockContext } from "../testHelper";
import { GraphApiClient } from "../../../src/clients/graphApi/client";
import { AADApplication } from "../../../src/clients/graphApi/types";

describe("GraphApiClient", () => {
  let sandbox: sinon.SinonSandbox;
  let mockAxios: {
    post: sinon.SinonStub;
    get: sinon.SinonStub;
    patch: sinon.SinonStub;
    delete: sinon.SinonStub;
    defaults: { headers: { common: Record<string, string> } };
    interceptors: {
      request: { use: sinon.SinonStub };
      response: { use: sinon.SinonStub };
    };
  };

  beforeEach(() => {
    sandbox = sinon.createSandbox();
    mockAxios = {
      post: sandbox.stub(),
      get: sandbox.stub(),
      patch: sandbox.stub(),
      delete: sandbox.stub(),
      defaults: { headers: { common: {} } },
      interceptors: {
        request: { use: sandbox.stub() },
        response: { use: sandbox.stub() },
      },
    };
    sandbox.stub(axios, "create").returns(mockAxios as any);
    sandbox.stub(retryModule, "sendWithRetry").callsFake(async (fn: any) => fn());
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("constructor", () => {
    it("sets Authorization and Content-Type headers", () => {
      const ctx = createMockContext();
      new GraphApiClient(ctx, "test-graph-token");
      expect(mockAxios.defaults.headers.common["Authorization"]).to.equal(
        "Bearer test-graph-token"
      );
      expect(mockAxios.defaults.headers.common["Content-Type"]).to.equal("application/json");
    });
  });

  describe("createAadApp", () => {
    it("creates an app with displayName", async () => {
      const app: AADApplication = { id: "obj-1", appId: "client-1", displayName: "My App" };
      mockAxios.post.resolves({ data: app, status: 200 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.createAadApp("My App");
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap().appId).to.equal("client-1");
      expect(mockAxios.post.calledOnce).to.be.true;
      expect(mockAxios.post.firstCall.args[0]).to.equal("/applications");
    });

    it("passes signInAudience and serviceManagementReference", async () => {
      mockAxios.post.resolves({ data: { id: "o1", appId: "c1" }, status: 200 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      await client.createAadApp("App", "AzureADMultipleOrgs", "svc-ref-123");
      const body = mockAxios.post.firstCall.args[1];
      expect(body.signInAudience).to.equal("AzureADMultipleOrgs");
      expect(body.serviceManagementReference).to.equal("svc-ref-123");
    });

    it("returns user error on 4xx", async () => {
      mockAxios.post.rejects({
        response: {
          status: 400,
          data: { error: { code: "BadRequest", message: "Invalid audience" } },
        },
      });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.createAadApp("App");
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().kind).to.equal("user");
    });

    it("returns system error on 5xx", async () => {
      mockAxios.post.rejects({
        response: { status: 500, data: { error: { code: "InternalError", message: "boom" } } },
      });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.createAadApp("App");
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().kind).to.equal("system");
    });
  });

  describe("generateClientSecret", () => {
    it("returns secretText on success", async () => {
      mockAxios.post.resolves({
        data: { secretText: "super-secret-value", keyId: "k1" },
        status: 200,
      });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.generateClientSecret("obj-123");
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.equal("super-secret-value");
    });

    it("returns error on empty secretText", async () => {
      mockAxios.post.resolves({ data: { secretText: "" }, status: 200 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.generateClientSecret("obj-123");
      expect(result.isErr()).to.be.true;
      expect(result._unsafeUnwrapErr().code).to.equal("EmptyClientSecret");
    });

    it("uses extra retries", async () => {
      mockAxios.post.resolves({ data: { secretText: "s" }, status: 200 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");
      await client.generateClientSecret("obj-123");

      const retryStub = retryModule.sendWithRetry as sinon.SinonStub;
      // second arg should be 5
      expect(retryStub.calledOnce).to.be.true;
      expect(retryStub.firstCall.args[1]).to.equal(5);
    });
  });

  describe("updateAadApp", () => {
    it("PATCHes the application", async () => {
      mockAxios.patch.resolves({ data: {}, status: 204 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.updateAadApp("obj-123", { displayName: "New Name" });
      expect(result.isOk()).to.be.true;
      expect(mockAxios.patch.firstCall.args[0]).to.equal("/applications/obj-123");
    });

    it("uses extra retries", async () => {
      mockAxios.patch.resolves({ data: {}, status: 204 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");
      await client.updateAadApp("obj-123", {});

      const retryStub = retryModule.sendWithRetry as sinon.SinonStub;
      expect(retryStub.firstCall.args[1]).to.equal(5);
    });
  });

  describe("getOwners", () => {
    it("returns owner list", async () => {
      mockAxios.get.resolves({
        data: { value: [{ id: "u1", displayName: "User 1" }] },
        status: 200,
      });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.getOwners("obj-123");
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.have.lengthOf(1);
    });

    it("returns empty array when value is missing", async () => {
      mockAxios.get.resolves({ data: {}, status: 200 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.getOwners("obj-123");
      expect(result.isOk()).to.be.true;
      expect(result._unsafeUnwrap()).to.deep.equal([]);
    });
  });

  describe("addOwner", () => {
    it("posts $ref link", async () => {
      mockAxios.post.resolves({ data: {}, status: 204 });
      const ctx = createMockContext();
      const client = new GraphApiClient(ctx, "tok");

      const result = await client.addOwner("obj-123", "user-456");
      expect(result.isOk()).to.be.true;
      expect(mockAxios.post.firstCall.args[0]).to.equal("/applications/obj-123/owners/$ref");
    });
  });
});
