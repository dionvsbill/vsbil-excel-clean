// exc/frontend-vite/src/components/PayButton.jsx
import React, { useState, useEffect } from "react";
import Toast from "./Toast";
import { supabase } from "../supabaseClient";

export default function PayButton({ currency = "GHS" }) {
  const apiBase = import.meta.env.VITE_API_URL || "http://localhost:8000";

  const [mode, setMode] = useState("one-time");
  const [loading, setLoading] = useState(false);
  const [showModal, setShowModal] = useState(false);
  const [email, setEmail] = useState(null);

  // Load the logged-in user's registration email
  useEffect(() => {
    const loadUserEmail = async () => {
      const { data } = await supabase.auth.getUser();
      setEmail(data?.user?.email || null);
    };
    loadUserEmail();
  }, []);

  const getAccessToken = async () => {
    const { data } = await supabase.auth.getSession();
    return data?.session?.access_token || null;
  };

  const payWithPaystack = async () => {
  if (!email) {
    Toast.error("No user email found");
    return;
  }
  setLoading(true);
  try {
    const token = await getAccessToken();

    // Correct minor units
    const amount =
      mode === "one-time" ? 150000 : 200;

    const res = await fetch(`${apiBase}/payments/init`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: token ? `Bearer ${token}` : "",
      },
      body: JSON.stringify({ email, amount}),
    });
    const data = await res.json();
    if (data.error) {
      Toast.error(data.error);
      setLoading(false);
      return;
    }

    if (data.data?.authorization_url) {
      window.open(data.data.authorization_url, "_blank");
      Toast.info("Redirecting to Paystack checkout...");
    } else {
      Toast.error("No authorization URL returned from backend");
    }

    setShowModal(false);
  } catch (err) {
    Toast.error("Init error: " + err.message);
  } finally {
    setLoading(false);
  }
};


  return (
    <div>
      <button className="primary" onClick={() => setShowModal(true)}>
        Subscribe
      </button>

      {showModal && (
        <div className="modal-overlay">
          <div className="modal">
            <h3>Select Subscription Option</h3>
            <label>
              <input
                type="radio"
                value="one-time"
                checked={mode === "one-time"}
                onChange={() => setMode("one-time")}
              />
              One‑time Payment (GHS 4999.99)
            </label>
            <label style={{ marginLeft: "1rem" }}>
              <input
                type="radio"
                value="monthly"
                checked={mode === "monthly"}
                onChange={() => setMode("monthly")}
              />
              Monthly Subscription (GHS 32.50)
            </label>

            <div style={{ marginTop: "1rem" }}>
              <button
                className="primary"
                onClick={payWithPaystack}
                disabled={loading}
              >
                {loading
                  ? "Processing..."
                  : mode === "one-time"
                  ? "Pay Once"
                  : "Subscribe Monthly"}
              </button>
              <button
                className="secondary"
                onClick={() => setShowModal(false)}
                style={{ marginLeft: "1rem" }}
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
