// Root of the static site. Outlook never opens this page — it opens /taskpane/.
export default function Home() {
  return (
    <main style={{ fontFamily: "Segoe UI, sans-serif", padding: 24 }}>
      <h1>BeInc</h1>
      <p>
        Outlook add-in. The task pane lives at <a href="/taskpane/">/taskpane/</a>.
      </p>
    </main>
  );
}
