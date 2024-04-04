#!/usr/bin/env node
import figlet from "figlet";
import { Command } from "commander";
import {
  PublicClientApplication,
  LogLevel,
  InteractiveRequest,
  Configuration,
} from "@azure/msal-node";
import fetch, { Headers } from "node-fetch";
import open from "open";
import { exit } from "process";
import { writeFile, mkdir } from "fs";

(async () => {
  const packageJson = await import("../package.json");

  console.log(
    figlet.textSync(packageJson.description, {
      font: "Small",
      width: 80,
      verticalLayout: "fitted",
      whitespaceBreak: true,
    })
  );

  const program = new Command();

  program
    .version(packageJson.version)
    .description(`${packageJson.description} by ${packageJson.author}`)
    .requiredOption("-u, --url <url>", "Dataverse Environment Url")
    .option("-o, --output <output>", "Output path", "./output")
    .option(
      "-c, --clientid <clientid>",
      "The application ID of your application",
      "51f81489-12ee-4a9e-aaae-a2591f45987d"
    )
    .option(
      "-a, --authority <authority>",
      "The authority URL for your application",
      "https://login.microsoftonline.com/common/"
    )
    .parse(process.argv);

  const opts = program.opts();
  const dvUrl = opts.url as string;
  const outputPath = opts.output as string;
  const authority = opts.authority as string;
  const clientid = opts.clientid as string;

  mkdir(outputPath, { recursive: true }, (err) => {
    if (err) {
      console.log(`Error creating output. ${err}`);
    }
  });

  const msalConfig: Configuration = {
    auth: {
      clientId: clientid,
      authority: authority,
    },
    system: {
      loggerOptions: {
        loggerCallback(loglevel: any, message: any, containsPii: any) {
          console.log(`${loglevel}|${message}|${containsPii}`);
        },
        piiLoggingEnabled: false,
        logLevel: LogLevel.Warning,
      },
    },
  };

  const openBrowser = async (url: string) => {
    open(url);
  };

  const loginRequest = {
    scopes: [`${dvUrl}/.default`],
  };

  const tokenRequest: InteractiveRequest = {
    scopes: [`${dvUrl}/.default`],
    openBrowser,
    successTemplate:
      "<h1>Successfully signed in!</h1> <p>You can close this window now.</p>",
    errorTemplate:
      "<h1>Oops! Something went wrong</h1> <p>Navigate back to the cli application and check the console for more information.</p>",
  };

  const pca = new PublicClientApplication(msalConfig);

  const auth = await pca.acquireTokenInteractive(tokenRequest);

  if (!auth) {
    console.log(`Authentication failed!`);
    exit(1);
  }

  let userid: string | undefined;
  const whoamiresponse = await fetch(`${dvUrl}/api/data/v9.0/WhoAmI`, {
    headers: { Authorization: `Bearer ${auth.accessToken}` },
  });
  if (whoamiresponse.ok) {
    const json = await whoamiresponse.json();
    userid = json.UserId as string;
    console.log(`Logged in UserId: ${userid}`);
  } else {
    console.log(`${whoamiresponse.status} ${whoamiresponse.statusText}`);
    exit(1);
  }

  const usersresponse = await fetch(
    `${dvUrl}/api/data/v9.2/systemusers?$select=systemuserid,azureactivedirectoryobjectid,fullname,caltype,userlicensetype,islicensed&$filter=(isdisabled eq false and azureactivedirectoryobjectid ne null)`,
    {
      headers: { Authorization: `Bearer ${auth.accessToken}` },
    }
  );
  let users: {
    systemuserid: string;
    fullname: string;
    azureactivedirectoryobjectid: string;
  }[];
  if (usersresponse.ok) {
    const json = await usersresponse.json();
    users = json.value;
    console.log(`Found ${users.length} users.`);
  } else {
    console.log(`${usersresponse.status} ${usersresponse.statusText}`);
    exit(1);
  }

  for (let index = 0; index < users.length; index++) {
    const user = users[index];
    console.log(`Extracting personal views for ${user.fullname}`);
    const viewsresponse = await fetch(`${dvUrl}/api/data/v9.2/userqueries`, {
      headers: new Headers({
        Authorization: `Bearer ${auth.accessToken}`,
        CallerObjectId: `${user.azureactivedirectoryobjectid}`,
        Prefer: "odata.include-annotations=*",
      }),
    });
    let views: any[] = [];
    if (viewsresponse.ok) {
      const json = await viewsresponse.json();
      views = json.value;
    } else {
      console.log(`${viewsresponse.status} ${viewsresponse.statusText}`);
      console.log(await viewsresponse.text());
    }

    for (let index = 0; index < views.length; index++) {
      const view = views[index];
      const filename = `${view.userqueryid}.json`;
      writeFile(
        `${outputPath}/viewdata_${filename}`,
        JSON.stringify(view),
        (err) => {
          if (err) {
            console.log(`Error saving file. ${err}`);
          }
        }
      );
    }

    console.log(`Extracted ${views.length} views for ${user.fullname}`);
  }

  exit();
})();
