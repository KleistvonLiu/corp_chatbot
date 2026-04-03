import "dotenv/config";
import { syncConfiguredKnowledgeBase } from "./lib/knowledge";
import { createApp } from "./app";

const port = Number(process.env.PORT ?? 3001);
const syncResult = await syncConfiguredKnowledgeBase();
if (syncResult.status === "synced") {
  console.log(
    `Fixed knowledge source synced: ${syncResult.knowledgeBaseId} (${syncResult.sourceCount} sources / ${syncResult.chunkCount} chunks)`
  );
} else if (syncResult.status === "failed") {
  console.error(`Fixed knowledge source sync failed: ${syncResult.error}`);
}
const app = await createApp();

app.listen(port, () => {
  console.log(`Server listening on http://localhost:${port}`);
});
