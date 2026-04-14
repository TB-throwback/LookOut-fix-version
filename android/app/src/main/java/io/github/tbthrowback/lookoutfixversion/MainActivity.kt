package io.github.tbthrowback.lookoutfixversion

import android.app.Activity
import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.webkit.ValueCallback
import android.webkit.WebChromeClient
import android.webkit.WebResourceRequest
import android.webkit.WebResourceResponse
import android.webkit.WebSettings
import android.webkit.WebView
import android.webkit.WebViewClient
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import org.json.JSONObject
import java.io.IOException
import java.net.URLConnection

class MainActivity : AppCompatActivity() {
	private lateinit var webView: WebView

	@Volatile
	private var pendingLaunch: LaunchFile? = null

	private var pageReady = false

	private var fileChooserCallback: ValueCallback<Array<Uri>>? = null

	private val fileChooserLauncher = registerForActivityResult(ActivityResultContracts.StartActivityForResult()) { result ->
		val data = result.data?.data ?: result.data?.clipData?.let { clipData ->
			if (clipData.itemCount > 0) clipData.getItemAt(0).uri else null
		}
		val uris = if (result.resultCode == Activity.RESULT_OK && data != null) arrayOf(data) else null
		fileChooserCallback?.onReceiveValue(uris)
		fileChooserCallback = null
	}

	override fun onCreate(savedInstanceState: Bundle?) {
		super.onCreate(savedInstanceState)
		setContentView(R.layout.activity_main)

		webView = findViewById(R.id.web_view)
		configureWebView()
		webView.loadUrl(BASE_URL)

		handleIntent(intent)
	}

	override fun onNewIntent(intent: Intent) {
		super.onNewIntent(intent)
		setIntent(intent)
		handleIntent(intent)
	}

	private fun configureWebView() {
		webView.settings.apply {
			javaScriptEnabled = true
			domStorageEnabled = true
			cacheMode = WebSettings.LOAD_NO_CACHE
			allowFileAccess = true
			allowContentAccess = true
		}

		webView.webChromeClient = object : WebChromeClient() {
			override fun onShowFileChooser(
				webView: WebView?,
				filePathCallback: ValueCallback<Array<Uri>>?,
				fileChooserParams: FileChooserParams?
			): Boolean {
				fileChooserCallback?.onReceiveValue(null)
				fileChooserCallback = filePathCallback
				val intent = fileChooserParams?.createIntent()
				if (intent != null) {
					fileChooserLauncher.launch(intent)
					return true
				}
				return false
			}
		}

		webView.webViewClient = object : WebViewClient() {
			override fun onPageFinished(view: WebView, url: String) {
				pageReady = true
				flushPendingLaunch()
			}

			override fun shouldInterceptRequest(
				view: WebView,
				request: WebResourceRequest,
			): WebResourceResponse? {
				val requestUrl = request.url
				if (requestUrl.host != HOST) {
					return null
				}

				return when (val path = requestUrl.path ?: "/") {
					"/input" -> inputResponse()
					"/", "/index.html" -> assetResponse("index.html")
					else -> assetResponse(path.removePrefix("/"))
				}
			}
		}
	}

	private fun handleIntent(intent: Intent?) {
		val launchFile = intent?.toLaunchFile() ?: return
		pendingLaunch = launchFile
		if (pageReady) {
			triggerExtraction(launchFile)
		}
	}

	private fun flushPendingLaunch() {
		val launchFile = pendingLaunch ?: return
		triggerExtraction(launchFile)
	}

	private fun triggerExtraction(launchFile: LaunchFile) {
		val script =
			"window.Lookout?.openFromAndroid(${JSONObject.quote(launchFile.name)}, ${JSONObject.quote(launchFile.mimeType)});"
		webView.post {
			webView.evaluateJavascript(script, null)
		}
	}

	private fun inputResponse(): WebResourceResponse? {
		val launchFile = pendingLaunch ?: return emptyResponse("No file was supplied.")
		val inputStream = contentResolver.openInputStream(launchFile.uri)
			?: return emptyResponse("Unable to open the supplied file.")
		return WebResourceResponse(
			launchFile.mimeType.ifBlank { "application/octet-stream" },
			null,
			inputStream,
		)
	}

	private fun assetResponse(assetName: String, mimeType: String = "application/octet-stream"): WebResourceResponse? {
		return try {
			val inputStream = assets.open(assetName)
			WebResourceResponse(guessMimeType(assetName, mimeType), "utf-8", inputStream)
		} catch (_: IOException) {
			null
		}
	}

	private fun emptyResponse(message: String): WebResourceResponse {
		val bytes = message.toByteArray(Charsets.UTF_8)
		return WebResourceResponse(
			"text/plain",
			"utf-8",
			bytes.inputStream(),
		)
	}

	private fun Intent.toLaunchFile(): LaunchFile? {
		val uri = data ?: return null
		val name = uri.displayName() ?: uri.lastPathSegment?.substringAfterLast('/') ?: "winmail.dat"
		val mimeType = contentResolver.getType(uri) ?: type ?: "application/octet-stream"
		return LaunchFile(uri, name, mimeType)
	}

	private fun Uri.displayName(): String? {
		contentResolver.query(this, arrayOf(android.provider.OpenableColumns.DISPLAY_NAME), null, null, null)
			?.use { cursor ->
				val index = cursor.getColumnIndex(android.provider.OpenableColumns.DISPLAY_NAME)
				if (index >= 0 && cursor.moveToFirst()) {
					return cursor.getString(index)
				}
			}
		return null
	}

	private fun guessMimeType(assetName: String, fallback: String = "application/octet-stream"): String {
		val guessed = URLConnection.guessContentTypeFromName(assetName)
		return guessed ?: when (assetName.substringAfterLast('.', "").lowercase()) {
			"html" -> "text/html"
			"css" -> "text/css"
			"js", "mjs" -> "text/javascript"
			"json", "webmanifest" -> "application/manifest+json"
			"png" -> "image/png"
			"jpg", "jpeg" -> "image/jpeg"
			"svg" -> "image/svg+xml"
			"avif" -> "image/avif"
			else -> fallback
		}
	}

	private data class LaunchFile(
		val uri: Uri,
		val name: String,
		val mimeType: String,
	)

	private companion object {
		private const val HOST = "localhost"
		private const val BASE_URL = "https://$HOST/index.html?android=1"
	}
}
