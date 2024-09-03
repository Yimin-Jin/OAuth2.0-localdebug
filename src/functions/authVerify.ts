import { HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import jwt from "jsonwebtoken";
import jwksClient from "jwks-rsa";

const client = jwksClient({
    jwksUri: `https://login.microsoftonline.com/${process.env.AAD_APP_TENANT_ID}/discovery/keys?appid=${process.env.TEAMS_APP_ID}`
});

function getKey(header, callback) {
    client.getSigningKey(header.kid, function (err, key) {
        if (err) {
            callback(err);
        } else {
            const signingKey = key.getPublicKey();
            callback(null, signingKey);
        }
    });
}

export async function authVerify(req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit | null> {
    const token = req.headers.get('authorization')?.split(' ')[1];
    if (!token) {
        return {
            status: 401,
            jsonBody: { error: "Unauthorized: No token provided" },
        };
    }

    try {
        await new Promise((resolve, reject) => {
            jwt.verify(token, getKey, (err, decoded) => {
                if (err) {
                    reject(err);
                } else {
                    context.log("Token decoded:", decoded);
                    resolve(decoded);
                }
            });
        });
        // access
        return null;
    } catch (err) {
        return {
            status: 401,
            jsonBody: { error: `Unauthorized: ${err.message}` },
        };
    }
}  
