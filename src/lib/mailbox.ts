/**
 * Reads the subject of the current mail item.
 * Compose mode exposes an async Office.Subject; read mode exposes a plain string.
 */
export async function getItemSubject(): Promise<string> {
  const item = Office.context.mailbox?.item;
  if (!item) throw new Error("No mail item is open.");

  const subject = item.subject as unknown as string | Office.Subject;
  if (typeof subject === "string") return subject;

  return new Promise<string>((resolve, reject) => {
    subject.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error.message));
    });
  });
}
