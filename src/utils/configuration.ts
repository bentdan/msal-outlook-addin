import { AppConfig } from "../react-app-env";

export async function getConfiguration(): Promise<AppConfig> {
  const response = await fetch('/app-config.json');
  return response.json();
}