"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const authHelpers_1 = require("../authHelpers");
function main(context, req) {
    return __awaiter(this, void 0, void 0, function* () {
        if (context)
            context.log("Starting Azure function!");
        let users = yield getAllUsers();
        let response = {
            status: 200,
            body: {
                users
            }
        };
        return response;
    });
}
exports.main = main;
;
function getAllUsers() {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield authHelpers_1.GraphClient();
        return client
            .api("/users")
            .get()
            .then((res) => {
            let users = res.value;
            return users;
        });
    });
}
// to test this locally in Visual Studio code, uncomment the following line
// when running in Azure functions, they'll call main() for you
// main(); 
//# sourceMappingURL=function.js.map