import {assert} from 'chai'

import { getClient, randomString } from "./testHelpers"
import { SchemaExtension, User } from '../microsoft-graph'

declare const describe, it;

let colorSchemaExtension:SchemaExtension = {
    id: `a830edad9050849NDA1_color`,
    description: "A schema that adds a color property to users",
    targetTypes: ["User"],
    properties: [{
        name: "color",
        type: "String"
    }]
}

describe('Schema Extensions', function() {
  this.timeout(10*1000);
  it('Use schema extensions to add a field to users', function() {
    return getClient().api("https://graph.microsoft.com/beta/schemaExtensions").post(colorSchemaExtension).then((json) => {
        let createdExtension = json as SchemaExtension;
        assert.equal(createdExtension.id, colorSchemaExtension.id);
        assert.equal(createdExtension.description, colorSchemaExtension.description);

        return Promise.resolve();
    });
  });

it('Updates the authenticated user with the extended property', function() {
    const sampleExtensionValues = {}

    sampleExtensionValues[colorSchemaExtension.id] = {
        color: `color-${randomString()}`
    };

    return getClient().api("https://graph.microsoft.com/beta/me").patch(sampleExtensionValues).then((json) => {
        return getClient().api(`https://graph.microsoft.com/beta/me?$select=${colorSchemaExtension.id}`).get().then((json) => {
            let user = json as User;

            assert(colorSchemaExtension.id in user, "User has schema in returned JSON");

            assert.equal(user[colorSchemaExtension.id].color, sampleExtensionValues[colorSchemaExtension.id].color);

            return Promise.resolve();

        });
    });
  });
});