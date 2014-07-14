package net.exkazuu.gmail_extractor

import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential
import com.google.api.client.googleapis.auth.oauth2.GoogleOAuthConstants
import com.google.api.client.http.javanet.NetHttpTransport
import com.google.api.client.json.jackson.JacksonFactory
import com.google.api.client.util.Base64
import com.google.api.services.gmail.Gmail
import java.io.BufferedReader
import java.io.File
import java.io.FileReader
import java.io.FileWriter
import java.io.InputStreamReader
import java.util.ArrayList
import java.util.Arrays
import org.supercsv.io.CsvBeanWriter
import org.supercsv.prefs.CsvPreference

public class GmailExtractor {

	// Check https://developers.google.com/gmail/api/auth/scopes for all available scopes
	static val SCOPE = "https://www.googleapis.com/auth/gmail.readonly";
	static val APP_NAME = "Gmail API Quickstart";

	// Email address of the user, or "me" can be used to represent the currently authorized user.
	static val USER = "me";

	// Path to the client_secret.json file downloaded from the Developer Console
	static val CLIENT_SECRET_PATH = "client_secret.json";

	val static headers = #["参加区分", "大学院名", "研究科", "専攻", "所属研究室", "学年", "氏名", "email", "電話番号", "GitHubアカウント名", "連絡事項"].
		toList

	static class SurveyResult {
		@Property String 参加区分 = ""
		@Property String 大学院名 = ""
		@Property String 研究科 = ""
		@Property String 専攻 = ""
		@Property String 所属研究室 = ""
		@Property String 学年 = ""
		@Property String 氏名 = ""
		@Property String email = ""
		@Property String 電話番号 = ""
		@Property String GitHubアカウント名 = ""
		@Property String 連絡事項 = ""
	}

	//static GoogleClientSecrets clientSecrets;
	def static void main(String[] args) {
		val httpTransport = new NetHttpTransport();
		val jsonFactory = new JacksonFactory();

		val clientSecrets = GoogleClientSecrets.load(jsonFactory, new FileReader(CLIENT_SECRET_PATH));

		// Allow user to authorize via url.
		val flow = new GoogleAuthorizationCodeFlow.Builder(httpTransport, jsonFactory, clientSecrets,
			Arrays.asList(SCOPE)).setAccessType("online").setApprovalPrompt("auto").build();

		val url = flow.newAuthorizationUrl().setRedirectUri(GoogleOAuthConstants.OOB_REDIRECT_URI).build();
		System.out.println(
			"Please open the following URL in your browser then type" + " the authorization code:\n" + url);

		// Read code entered by user.
		val br = new BufferedReader(new InputStreamReader(System.in));
		val code = br.readLine();

		// Generate Credential using retrieved code.
		val response = flow.newTokenRequest(code).setRedirectUri(GoogleOAuthConstants.OOB_REDIRECT_URI).execute();
		val credential = new GoogleCredential().setFromTokenResponse(response);

		// Create a new authorized Gmail API client
		val service = new Gmail.Builder(httpTransport, jsonFactory, credential).setApplicationName(APP_NAME).build();

		// Retrieve a page of Threads; max of 100 by default.
		val messageResponse = service.users().messages().list(USER).setQ("お申込みがありました").setMaxResults(200L).execute()
		val messages = messageResponse.messages

		// Print ID of each Thread.
		val surveyResults = new ArrayList<GmailExtractor.SurveyResult>()
		for (message : messages) {
			val msg = service.users.messages.get(USER, message.id).setFormat("full").execute
			val body = new String(Base64.decodeBase64(msg.payload.body.data))
			if (body.startsWith("サイトよりお申込みがありました。")) {
				val surveyText = body.substring(0, body.lastIndexOf("--"))

				val part = surveyText.indexOf("参加区分")
				val univ = surveyText.indexOf("大学院名", part + 1)
				val div = surveyText.indexOf("研究科", univ + 1)
				val dep = surveyText.indexOf("専攻", div + 1)
				val lab = surveyText.indexOf("所属研究室", dep + 1)
				val year = surveyText.indexOf("学年", lab + 1)
				val name = surveyText.indexOf("氏名", year + 1)
				val email = surveyText.indexOf("email", name + 1)
				val tel = surveyText.indexOf("電話番号", email + 1)
				val github = surveyText.indexOf("GitHubアカウント名", tel + 1)
				val other = surveyText.indexOf("連絡事項", github + 1)

				val ret = new GmailExtractor.SurveyResult()
				ret.参加区分 = surveyText.substring(part + "参加区分".length, univ).trim
				ret.大学院名 = surveyText.substring(univ + "大学院名".length, div).trim
				ret.研究科 = surveyText.substring(div + "研究科".length, dep).trim
				ret.専攻 = surveyText.substring(dep + "専攻".length, lab).trim
				ret.所属研究室 = surveyText.substring(lab + "所属研究室".length, year).trim
				ret.学年 = surveyText.substring(year + "学年".length, name).trim
				ret.氏名 = surveyText.substring(name + "氏名".length, email).trim
				ret.email = surveyText.substring(email + "email".length, tel).trim
				ret.電話番号 = surveyText.substring(tel + "電話番号".length, github).trim
				if (other > 0) {
					ret.gitHubアカウント名 = surveyText.substring(github + "GitHubアカウント名".length, other).trim
					ret.連絡事項 = surveyText.substring(other + "連絡事項".length).trim
				} else {
					ret.gitHubアカウント名 = surveyText.substring(github + "GitHubアカウント名".length).trim
				}
				surveyResults.add(ret)
			}
			System.out.println(surveyResults.size + " / " + messages.size)
		}
		write(new File("result.csv"), surveyResults)
	}

	def static write(File file, Iterable<GmailExtractor.SurveyResult> results) {
		val writer = new FileWriter(file)
		val csvWriter = new CsvBeanWriter(writer, CsvPreference.STANDARD_PREFERENCE)
		csvWriter.writeHeader(headers)
		for (result : results) {
			csvWriter.write(result, headers)
		}
		csvWriter.close
		writer.close
	}
}
