const displayMessages = [];
let currentTenantId = "";
let context = {};

// Call the initialize API first
microsoftTeams.app.initialize().then(function () {
  microsoftTeams.app.getContext().then(function (ctx) {
    // Save the context object
    context = ctx;
    currentTenantId = ctx.user.tenant.id;
    showMessage(context);
  });
});

function showMessage(message) {
  if (message) {
    const msg =
      typeof message === "string" ? message : JSON.stringify(message, null, 2);
    displayMessages.push(msg);
    // keep only 20 messages
    if (displayMessages.length > 20) {
      displayMessages.shift();
    }
    document.getElementById("user-details").innerHTML = displayMessages
      .reverse()
      .join("<br>");
    displayMessages.reverse();
  }
}
async function getAuthToken() {
  return microsoftTeams.authentication.getAuthToken({
    tenantId: currentTenantId, //"9ff54c41-79f0-4837-a141-2e87ebfbabdf", //currentTenantId,
  });
}
async function userTokenGetMe() {
  let token;
  try {
    token = await getAuthToken();
  } catch (error) {
    showMessage(error.message);
    console.log(error);
    return;
  }

  fetch("/user-token/me", {
    method: "POST",
    body: JSON.stringify({ token, tid: currentTenantId }),
    headers: {
      "Content-Type": "application/json",
    },
  })
    .then((result) => {
      return result.json();
    })
    .then((user) => {
      showMessage(user);
    })
    .catch((error) => {
      showMessage(error);
    });
}

async function userTokenGetChannelMember() {
  const teamId = context.channel.ownerGroupId;
  const channelId = context.channel.id;
  const token = await getAuthToken();
  fetch(`/user-token/teams/${teamId}/channels/${channelId}/members`, {
    method: "POST",
    body: JSON.stringify({ token, tid: currentTenantId }),
    headers: {
      "Content-Type": "application/json",
    },
  })
    .then((result) => {
      return result.json();
    })
    .then((user) => {
      showMessage(user);
    })
    .catch((error) => {
      showMessage(error);
    });
}

async function appTokenGetChannelMember() {
  const teamId = context.channel.ownerGroupId;
  const channelId = context.channel.id;
  fetch(`/app-token/teams/${teamId}/channels/${channelId}/members`, {
    method: "POST",
    body: JSON.stringify({}),
    headers: {
      "Content-Type": "application/json",
    },
  })
    .then((result) => {
      return result.json();
    })
    .then((user) => {
      showMessage(user);
    })
    .catch((error) => {
      showMessage(error);
    });
}
