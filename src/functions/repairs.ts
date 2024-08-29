import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";  
import repairRecords from "../repairsData.json";  
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
  
/**   
 * This function handles the HTTP request and returns the repair information.  
 *   
 * @param {HttpRequest} req - The HTTP request.  
 * @param {InvocationContext} context - The Azure Functions context object.  
 * @returns {Promise<Response>} - A promise that resolves with the HTTP response containing the repair information.  
 */  
export async function repairs(req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {  
  context.log("HTTP trigger function processed a request.");  
  
  // Initialize response  
  const res: HttpResponseInit = {  
    status: 200,  
    jsonBody: {  
      results: [],  
    },  
  };  
  
  // Get the token from the request headers  
  const token = req.headers.get('authorization')?.split(' ')[1].toString();
  
  if (!token) {  
    return {  
      status: 401,  
      jsonBody: { error: "Unauthorized: No token provided" },  
    };  
  }  
  
  try {  
    // Verify the token using the getKey callback  
    const decoded = await new Promise((resolve, reject) => {  
      jwt.verify(token, getKey, (err, decoded) => {  
        if (err) {  
          reject(err);  
        } else {  
          resolve(decoded);  
        }  
      });  
    });  
  
    // Log the decoded payload  
    context.log("Token decoded:", decoded);  
  
    // Get the assignedTo query parameter  
    const assignedTo = req.query.get("assignedTo");  
  
    // If the assignedTo query parameter is not provided, return all repair records  
    if (!assignedTo) {  
      res.jsonBody.results = repairRecords;  
      return res;  
    }  
  
    // Filter the repair information by the assignedTo query parameter  
    const repairs = repairRecords.filter((item) => {  
      const fullName = item.assignedTo.toLowerCase();  
      const query = assignedTo.trim().toLowerCase();  
      const [firstName, lastName] = fullName.split(" ");  
      return fullName === query || firstName === query || lastName === query;  
    });  
  
    // Return filtered repair records, or an empty array if no records were found  
    res.jsonBody.results = repairs ?? [];  
    return res;  
  } catch (err) {  
    // Handle JWT verification errors  
    if (err.name === 'TokenExpiredError') {  
      return {  
        status: 401,  
        jsonBody: { error: "Unauthorized: Token expired" },  
      };  
    } else if (err.name === 'JsonWebTokenError') {  
      return {  
        status: 401,  
        jsonBody: { error: `Unauthorized: ${err.message}` },  
      };  
    } else {  
      return {  
        status: 500,  
        jsonBody: { error: "Internal Server Error" },  
      };  
    }  
  }  
}  
   
  
app.http("repairs", {  
  methods: ["GET"],  
  authLevel: "anonymous",  
  handler: repairs,  
});  
