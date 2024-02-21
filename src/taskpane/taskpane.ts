/* global Office */

import axios from "axios";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Ensure DOM content is fully loaded
    document.getElementById("app-body").style.display = "block";
    document.getElementById("loginForm")?.addEventListener("submit", login);
    document.getElementById("signoutButton")?.addEventListener("click", signoutUser);
    document.getElementById("signupButton")?.addEventListener("click", signupUser);
    document.getElementById("signinButton")?.addEventListener("click", autoLoginUser);
    document.getElementById("searchInput")?.addEventListener("keypress", function (event) {
      if (event.key === "Enter") {
        searchGifs();
      }
    });
  }
});

async function signoutUser() {
  try {
    Office.context.roamingSettings.remove("accessToken");
    Office.context.roamingSettings.saveAsync(function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        document.getElementById("app-body").style.display = "block";
        document.getElementById("search-form").style.display = "none";
        document.getElementById("loading-spinner").style.display = "none";
        document.getElementById("divider").style.display = "none";
        document.getElementById("gifs-container").style.display = "none";
        document.getElementById("manual-login-form").style.display = "none";
        document.getElementById("signoutButton").style.display = "none";
        document.getElementById("landing-page").style.display = "flex";

        console.log("Signed out successfully.");
      } else {
        console.error("Error saving settings: " + asyncResult.error.message);
      }
    });
  } catch (error) {
    console.error("Error during sign-out:", error);
  }
}

let allGifs = [];
let isManualLoginInProgress = false;

async function autoLoginUser() {
  if (isManualLoginInProgress) {
    return;
  }
  const userEmail = Office.context.mailbox.userProfile.emailAddress;
  document.getElementById("landing-page").style.display = "none";
  try {
    const response = await axios.post("https://gift-server-eu-1.azurewebsites.net/login_with_email", {
      email: userEmail,
    });

    const result = response.data;
    console.log('result', result);
    if (result.status === "Login successful") {
      const accessToken = result.access_token;
      Office.context.roamingSettings.set("accessToken", accessToken);
      document.getElementById("signoutButton").style.display = "flex";
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
  const email = (document.getElementById("email") as HTMLInputElement).value;
  const password = (document.getElementById("password") as HTMLInputElement).value;

  try {
    const response = await axios.post(`https://gift-server-eu-1.azurewebsites.net/signin`, {
      email,
      password,
    });
    console.log("Sign in successful:", response.data);

    Office.context.roamingSettings.set("accessToken", response.data.access_token);
    document.getElementById("signoutButton").style.display = "flex";
    console.log("response.data.accessToken", response.data.access_token);
    console.log("response.data.accessToken", Office.context.roamingSettings.get("accessToken"));
    Office.context.roamingSettings.saveAsync(function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Error saving settings: " + asyncResult.error.message);
      } else {
        document.getElementById("manual-login-form").style.display = "none";
        fetchAndDisplayUserGifs();
      }
    });
  } catch (error) {
    console.error("Error signing in:", error);
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
    document.getElementById("divider").style.display = "flex";
    const gifs = response.data.data; // Adjust based on actual response structure
    allGifs = gifs || [];
    console.log("response", response);
    if (allGifs.length === 0) {
      displayCreateGifButton();
    } else {
      displayGifs(allGifs);
    }
  } catch (error) {
    console.error("Error fetching user gifs:", error);
  }
}

function searchGifs() {
  const searchTerm = (document.getElementById("searchInput") as HTMLInputElement).value.toLowerCase();
  const filteredGifs = allGifs.filter((gif) => gif.name.toLowerCase().includes(searchTerm));
  displayGifs(filteredGifs);
}

function calculateBase64FileSize(base64String) {
  // Remove MIME type prefix if present
  const base64Data = base64String.substring(base64String.indexOf(",") + 1);
  // Remove all spaces/newlines if any
  const cleanBase64Data = base64Data.replace(/\s+/g, "");
  // n is the length of the base64 encoded string
  const n = cleanBase64Data.length;
  // y is the padding count indicated by '=' characters at the end of the string
  let y = 0;
  if (cleanBase64Data.endsWith("==")) {
    y = 2;
  } else if (cleanBase64Data.endsWith("=")) {
    y = 1;
  }
  // Calculate the file size in bytes
  const fileSizeInBytes = n * (3 / 4) - y;
  return fileSizeInBytes;
}

function insertGifIntoCurrentEmail(sourceUrl, exampleEmail, base64, resourceId) {
  const maxSize = 1250000;
  const fileSizeInBytes = calculateBase64FileSize(base64);
  console.log("length", fileSizeInBytes);

  if (base64.length > maxSize) {
    document.getElementById("custom-modal").style.display = "block";
    document.getElementById("close-modal").onclick = function () {
      document.getElementById("custom-modal").style.display = "none";
    };
    return;
  } else {
    Office.context.mailbox.item.body.getAsync(
      "html",
      { asyncContext: "This is passed to the callback" },
      function callback(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const existingBody = result.value;
          const formattedExampleEmail = exampleEmail.replace(/\n\n/g, "<br><br>") || "";

          const gifHtml = `
            ${formattedExampleEmail}
            <table style="max-width: 300px;">
              <tr>
                <td style="border:none;">
                  <a href="https://giveagif-t.com/validate-gif?gif_id=${resourceId}" target="_blank">
                    <img src="data:image/gif;base64,${base64}" alt="GIF" style="width: 100%; height: auto;"/>
                  </a>
                </td>
              </tr>
            </table>
          `;

          const updatedBody = existingBody + gifHtml;
          Office.context.mailbox.item.body.setAsync(
            updatedBody,
            { coercionType: Office.CoercionType.Html, asyncContext: "This is passed to the callback" },
            function (result) {
              console.log("result", result);
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("GIF inserted successfully!");
              } else {
                document.getElementById("custom-modal").style.display = "block";
                document.getElementById("close-modal").onclick = function () {
                  document.getElementById("custom-modal").style.display = "none";
                };
              }
            }
          );
        } else {
          document.getElementById("custom-modal").style.display = "block";
          document.getElementById("close-modal").onclick = function () {
            document.getElementById("custom-modal").style.display = "none";
          };
        }
      }
    );
  }
}

async function signupUser() {
  try {
    window.open("https://giveagif-t.com/signup", "_blank");
  } catch (error) {
    console.error("Error during auto-login:", error);
  }
}

function displayCreateGifButton() {
  const container = document.getElementById("gifs-container");
  container.style.display = "flex";
  container.innerHTML = "";

  const button = document.createElement("button");
  button.innerText = "Create GIFs";
  button.style.cursor = "pointer";
  button.style.display = "flex";
  button.onclick = function () {
    window.open("https://giveagif-t.com", "_blank");
  };

  // Styling the button (Optional)
  button.style.backgroundColor = "#F4149B";
  button.style.color = "white";
  button.style.border = "none";
  button.style.padding = "10px 20px";
  button.style.width = "150px";
  button.style.borderRadius = "50px";
  button.style.fontSize = "15px";

  container.appendChild(button);
}

function displayGifs(gifs) {
  const container = document.getElementById("gifs-container");
  container.style.display = "flex";

  if (!container) {
    console.error("GIFs container not found");
    return;
  }

  container.innerHTML = "";

  gifs?.forEach((gif) => {
    const gifContainer = document.createElement("div");
    const img = document.createElement("img");
    const name = document.createElement("span");
    gifContainer.style.height = "140px";

    gifContainer.style.overflow = "hidden";
    img.src = gif.url;
    img.alt = "User GIF";
    img.style.width = "100px";
    img.style.height = "100px";
    img.style.cursor = "pointer";
    name.style.width = "80px";
    name.style.margin = "auto";
    name.style.whiteSpace = "no-wrap";
    name.style.overflow = "hidden";
    name.style.fontSize = "12px";
    name.style.color = "#fff";
    name.style.fontFamily = "Staatliches";
    img.addEventListener("click", () =>
      insertGifIntoCurrentEmail(gif.source || "https://gif-t.io", gif.example_email || "", gif.base64, gif.resourceId)
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
