/* global Office */

import axios from "axios";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("loginForm").addEventListener("submit", login);
    document.getElementById("searchInput").addEventListener("keypress", function (event) {
      if (event.key === "Enter") {
        searchGifs();
      }
    });
    document.getElementById("createGifsButton").addEventListener("click", function () {
      window.open("https://giveagif-t.com/gif-library", "_blank");
    });
    setTimeout(() => {
      autoLoginUser();
    }, 3000);
  }
});

let allGifs = [];
let isManualLoginInProgress = false;

async function autoLoginUser() {
  if (isManualLoginInProgress) {
    return;
  }
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
      displayManualLoginForm();
    }
  } catch (error) {
    console.error("Error during auto-login:", error);
    displayManualLoginForm();
  }
}

function displayManualLoginForm() {
  document.getElementById("manual-login-form").style.display = "block";
}

export async function login(event) {
  event.preventDefault();
  isManualLoginInProgress = true;
  const email = (document.getElementById("emailManual") as HTMLInputElement).value;
  const password = (document.getElementById("passwordManual") as HTMLInputElement).value;

  try {
    const response = await axios.post(`https://gift-server-eu-1.azurewebsites.net/signin`, {
      email,
      password,
    });
    console.log("Sign in successful:", response.data);

    Office.context.roamingSettings.set("accessToken", response.data.access_token);
    console.log("response.data.accessToken", response.data.access_token);
    console.log("response.data.accessToken", Office.context.roamingSettings.get("accessToken"));
    Office.context.roamingSettings.saveAsync(function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Error signing in:" + asyncResult.error.message);
      } else {
        document.getElementById("manual-login-form").style.display = "none";
        document.getElementById("error").style.display = "none";
        fetchAndDisplayUserGifs();
      }
    });
  } catch (error) {
    console.error("Error signing in:", error);
    document.getElementById("error").style.display = "flex";
  }
  isManualLoginInProgress = false;
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
    document.getElementById("search-form").style.display = "flex";
    const gifs = response.data.data;
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

function insertGifIntoCurrentEmail(gifUrl, sourceUrl, exampleEmail) {
  Office.context.mailbox.item.body.getAsync(
    "html",
    { asyncContext: "This is passed to the callback" },
    function callback(result) {
      // Check for success
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // The existing body content
        const existingBody = result.value;
        const verifiedWatermarkUrl =
          "https://gift-general-resources.s3.eu-north-1.amazonaws.com/verified_by_gift_2.png";
        const formattedExampleEmail = exampleEmail.replace(/\n\n/g, "<br><br>") || "";

        // Construct the HTML content to insert
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

        // Combine with existing body and update the email body
        const updatedBody = existingBody + gifHtml;
        Office.context.mailbox.item.body.setAsync(
          updatedBody,
          { coercionType: Office.CoercionType.Html, asyncContext: "This is passed to the callback" },
          function callback(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              console.log("GIF inserted successfully!");
            } else {
              console.error("Failed to insert GIF:", result.error);
            }
          }
        );
      } else {
        console.error("Failed to get email body:", result.error);
      }
    }
  );
}

function displayGifs(gifs) {
  const container = document.getElementById("gifs-container");
  if (!container) {
    console.error("GIFs container not found");
    return;
  }

  container.innerHTML = "";

  if (gifs.length === 0) {
    document.getElementById("no-gifs").style.display = "flex";
    document.getElementById("search-form").style.display = "none";
  } else {
    gifs.forEach((gif) => {
      const gifContainer = document.createElement("div");
      const img = document.createElement("img");
      const name = document.createElement("span");
      gifContainer.style.overflow = "hidden";
      img.src = gif.url;
      img.alt = "User GIF";
      img.style.width = "130px";
      img.style.height = "130px";
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
