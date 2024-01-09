/* global Office */

import axios from "axios";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("loginButton").addEventListener("click", login);
    document.getElementById("searchInput").addEventListener("keypress", function (event) {
      if (event.key === "Enter") {
        searchGifs();
      }
    });
    autoLoginUser();
  }
});

let allGifs = [];

async function autoLoginUser() {
  const userEmail = Office.context.mailbox.userProfile.emailAddress;
  console.log("userEmail", userEmail);

  try {
    const response = await axios.post("https://gift-server-eu-1.azurewebsites.net/login_with_email", {
      email: userEmail,
    });

    const result = response.data;
    if (result.status === "Login successful") {
      const accessToken = result.access_token;
      Office.context.roamingSettings.set("accessToken", accessToken);
      Office.context.roamingSettings.saveAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Error saving access token:", asyncResult.error.message);
        } else {
          fetchAndDisplayUserGifs();
        }
      });
    } else {
      console.log("User login failed or user does not exist");
      document.getElementById("login-form").style.display = "flex";
    }
  } catch (error) {
    console.error("Error during auto-login:", error);
  }
}

export async function fetchAndDisplayUserGifs() {
  const accessToken = Office.context.roamingSettings.get("accessToken");
  if (!accessToken) {
    console.error("Access token is not available");
    return;
  }

  try {
    const loadingSpinner = document.getElementById("loading-spinner");
    if (loadingSpinner) loadingSpinner.style.display = "flex";
    const response = await axios.get(`https://gift-server-eu-1.azurewebsites.net/fetch_user_gifs`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (loadingSpinner) loadingSpinner.style.display = "none";
    const gifs = response.data.data; // Adjust based on actual response structure
    allGifs = gifs || [];
    console.log("response", response);
    displayGifs(allGifs);
  } catch (error) {
    console.error("Error fetching user gifs:", error);
  }
}

function searchGifs() {
  const searchTerm = (document.getElementById("searchInput") as HTMLInputElement).value.toLowerCase();
  const filteredGifs = allGifs.filter((gif) => gif.name.toLowerCase().includes(searchTerm));
  displayGifs(filteredGifs);
}

function createDraftWithGif(gifUrl: string, sourceUrl: string, exampleEmail: string) {
  const verifiedWatermarkUrl = "https://gift-general-resources.s3.eu-north-1.amazonaws.com/verified_by_gift_2.png";
  const formattedExampleEmail = exampleEmail.replace(/\n\n/g, "<br><br>") || "";

  const htmlBody = `
    ${formattedExampleEmail}
    <table style="width: 200px; margin-bottom: 20px;">
      <tr>
        <td style="border:none;">
          <a href="${sourceUrl}" target="_blank">
            <img src="${gifUrl}" alt="GIF" style="width: 100%; height: auto;"/>
          </a>
        </td>
      </tr>
      <tr>
        <td style="border:none;">
          <img src="${verifiedWatermarkUrl}" alt="Verified" style="width: 100%; height: auto;"/>
        </td>
      </tr>
    </table>
  `;

  Office.context.mailbox.displayNewMessageForm({
    htmlBody: htmlBody,
    // Add other properties like subject, to recipients, cc recipients, etc., as needed
  });
}

function insertGifIntoCurrentEmail(gifUrl: string, sourceUrl: string, exampleEmail: string) {
  // Get the current item (email) the user is working on
  const item = Office.context.mailbox.item as Office.MessageCompose;

  // Construct the HTML content to insert
  const verifiedWatermarkUrl = "https://gift-general-resources.s3.eu-north-1.amazonaws.com/verified_by_gift_2.png";
  const formattedExampleEmail = exampleEmail.replace(/\n\n/g, "<br><br>") || "";
  const gifHtml = `
    ${formattedExampleEmail}
    <table style="width: 200px; margin-bottom: 20px;">
      <tr>
        <td style="border:none;">
          <a href="${sourceUrl}" target="_blank">
            <img src="${gifUrl}" alt="GIF" style="width: 100%; height: auto;"/>
          </a>
        </td>
      </tr>
      <tr>
        <td style="border:none;">
          <img src="${verifiedWatermarkUrl}" alt="Verified" style="width: 100%; height: auto;"/>
        </td>
      </tr>
    </table>
  `;

  // Insert the GIF HTML into the body of the email
  item.body.setSelectedDataAsync(gifHtml, { coercionType: Office.CoercionType.Html }, (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error("Error inserting GIF: " + result.error.message);
    }
  });
}

function displayGifs(gifs) {
  const container = document.getElementById("gifs-container");
  if (!container) {
    console.error("GIFs container not found");
    return;
  }

  container.innerHTML = "";

  gifs?.forEach((gif) => {
    const gifContainer = document.createElement("div"); // Container for each GIF and its name
    const img = document.createElement("img");
    const name = document.createElement("span");
    gifContainer.style.height = "150px";
    gifContainer.style.overflow = "hidden";
    img.src = gif.url;
    img.alt = "User GIF";
    img.style.width = "120px";
    img.style.height = "120px";
    img.style.cursor = "pointer";
    name.style.width = "120px";
    name.style.overflow = "hidden";
    name.style.color = "#fff";
    name.style.fontFamily = "Staatliches";
    img.addEventListener("click", () =>
      insertGifIntoCurrentEmail(gif.url, gif.source || "https://gif-t.io", gif.example_email || "")
    );

    name.textContent = gif.name;
    name.style.display = "block";
    name.style.textAlign = "center";

    gifContainer.appendChild(img);
    gifContainer.appendChild(name);

    container.appendChild(gifContainer);
  });
}
export async function run() {
  try {
    const item = Office.context.mailbox.item;
    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
    const userEmail = await getUserEmail();
    // await loginUser(userEmail);
    // Add more functionality as needed
  } catch (error) {
    console.error("Error:", error);
    // Handle errors appropriately
  }
}

function getUserEmail(): string {
  return Office.context.mailbox.userProfile.emailAddress;
}

export async function login() {
  const email = (document.getElementById("email") as HTMLInputElement).value;
  const password = (document.getElementById("password") as HTMLInputElement).value;

  try {
    const response = await axios.post(`https://gift-server-eu-1.azurewebsites.net/signin`, {
      email,
      password,
    });
    console.log("Sign in successful:", response.data);

    // Store the access token
    Office.context.roamingSettings.set("accessToken", response.data.access_token);
    console.log("response.data.accessToken", response.data.access_token);
    console.log("response.data.accessToken", Office.context.roamingSettings.get("accessToken"));
    Office.context.roamingSettings.saveAsync(function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Error saving settings: " + asyncResult.error.message);
      } else {
        fetchAndDisplayUserGifs();
        console.log("Settings saved");
      }
    });
    // Handle successful sign-in
  } catch (error) {
    console.error("Error signing in:", error);
    // Handle errors
  }
}

// Additional functions for your add-in as needed
