import { HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import jwt from "jsonwebtoken";
import jwksClient from "jwks-rsa";

// Initialize the JWKS client with the JWKS URI  
const client = jwksClient({
    jwksUri: `https://login.microsoftonline.com/${process.env.AAD_APP_TENANT_ID}/discovery/keys?appid=${process.env.TEAMS_APP_ID}`
});

/**  
 * Retrieves the signing key for JWT verification.  
 *  
 * @param header - The JWT header.  
 * @param callback - The callback function to return the signing key.  
 */
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

/**  
 * Middleware function to handle authorization using JWT.  
 *  
 * @param {HttpRequest} req - The HTTP request.  
 * @param {InvocationContext} context - The Azure Functions context object.  
 * @returns {Promise<HttpResponseInit | null>} - A promise that resolves with the HTTP response if unauthorized, or null if authorized.  
 */
export async function authMiddleware(req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit | null> {
    
    // Get the token from the request headers
    const token = req.headers.get('authorization')?.split(' ')[1];
    if (!token) {
        return {
            status: 401,
            jsonBody: { error: "Unauthorized: No token provided" },
        };
    }

    try {
        // Verify the token using the getKey callback
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
        // Indicates authorization is successful
        return null;
    } catch (err) {
        // Handle JWT verification errors
        return {
            status: 401,
            jsonBody: { error: `Unauthorized: ${err.message}` },
        };
    }
}  
