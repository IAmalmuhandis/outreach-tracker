const signinForm = document.getElementById("signinForm");
const signupForm = document.getElementById("signupForm");
const showSigninBtn = document.getElementById("showSigninBtn");
const showSignupBtn = document.getElementById("showSignupBtn");

async function api(path, options = {}) {
  const headers = { ...(options.headers || {}) };
  if (!(options.body instanceof FormData) && !headers["Content-Type"]) {
    headers["Content-Type"] = "application/json";
  }
  const response = await fetch(path, { headers, ...options });
  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.message || "Request failed");
  }
  if (response.status === 204) return null;
  return response.json();
}

function showSignin() {
  signinForm.hidden = false;
  signupForm.hidden = true;
}

function showSignup() {
  signinForm.hidden = true;
  signupForm.hidden = false;
}

async function checkExistingSession() {
  try {
    await api("/api/auth/session");
    window.location.replace("/");
  } catch {
    showSignin();
  }
}

showSigninBtn.addEventListener("click", showSignin);
showSignupBtn.addEventListener("click", showSignup);

signinForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  try {
    const payload = Object.fromEntries(new FormData(signinForm).entries());
    await api("/api/auth/signin", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    window.location.replace("/");
  } catch (error) {
    alert(`Sign in failed: ${error.message}`);
  }
});

signupForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  try {
    const payload = Object.fromEntries(new FormData(signupForm).entries());
    await api("/api/auth/signup", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    window.location.replace("/");
  } catch (error) {
    alert(`Sign up failed: ${error.message}`);
  }
});

checkExistingSession();
