package io.github.tbthrowback.lookoutfixversion

import android.app.Activity
import android.content.ContentValues
import android.content.Intent
import android.util.Base64
import android.os.Build
import android.os.Environment
import android.net.Uri
import android.os.Bundle
import android.provider.MediaStore
import android.webkit.ValueCallback
import android.webkit.WebChromeClient
import android.webkit.JavascriptInterface
import android.webkit.WebResourceRequest
import android.webkit.WebResourceResponse
import android.webkit.WebSettings
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.Toast
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

	private var pendingSaveRequest: PendingSaveRequest? = null

	private val fileChooserLauncher = registerForActivityResult(ActivityResultContracts.StartActivityForResult()) { result ->
		val data = result.data?.data ?: result.data?.clipData?.let { clipData ->
			if (clipData.itemCount > 0) clipData.getItemAt(0).uri else null
		}
		val uris = if (result.resultCode == Activity.RESULT_OK && data != null) arrayOf(data) else null
		fileChooserCallback?.onReceiveValue(uris)
		fileChooserCallback = null
	}

	private val createDocumentLauncher = registerForActivityResult(ActivityResultContracts.CreateDocument("*/*")) { outputUri ->
		val request = pendingSaveRequest
		pendingSaveRequest = null
		if (request == null) {
			return@registerForActivityResult
		}

		if (outputUri == null) {
			Toast.makeText(this, "Download canceled.", Toast.LENGTH_SHORT).show()
			return@registerForActivityResult
		}

		try {
			contentResolver.openOutputStream(outputUri)?.use { output ->
				output.write(request.bytes)
			}
			Toast.makeText(this, "Saved ${request.fileName}", Toast.LENGTH_SHORT).show()
			openDownloadedFile(outputUri, request.mimeType)
		} catch (_: IOException) {
			Toast.makeText(this, "Unable to save ${request.fileName}", Toast.LENGTH_LONG).show()
		}
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

		webView.addJavascriptInterface(AndroidBridge(), "AndroidBridge")

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
		val uri = when (action) {
			Intent.ACTION_SEND -> extraStreamUri() ?: data
			Intent.ACTION_SEND_MULTIPLE -> {
				val streams = extraStreamUris()
				if (streams.isNotEmpty()) streams.first() else data
			}
			else -> data ?: extraStreamUri()
		} ?: clipData?.let { clip ->
			if (clip.itemCount > 0) clip.getItemAt(0).uri else null
		} ?: return null
		val name = uri.displayName() ?: uri.lastPathSegment?.substringAfterLast('/') ?: "winmail.dat"
		val mimeType = contentResolver.getType(uri) ?: type ?: "application/octet-stream"
		return LaunchFile(uri, name, mimeType)
	}

	private fun Intent.extraStreamUri(): Uri? {
		return if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU) {
			getParcelableExtra(Intent.EXTRA_STREAM, Uri::class.java)
		} else {
			@Suppress("DEPRECATION")
			getParcelableExtra(Intent.EXTRA_STREAM)
		}
	}

	private fun Intent.extraStreamUris(): List<Uri> {
		return if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU) {
			getParcelableArrayListExtra(Intent.EXTRA_STREAM, Uri::class.java).orEmpty()
		} else {
			@Suppress("DEPRECATION")
			getParcelableArrayListExtra<Uri>(Intent.EXTRA_STREAM).orEmpty()
		}
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

	private data class PendingSaveRequest(
		val fileName: String,
		val mimeType: String,
		val bytes: ByteArray,
	)

	private inner class AndroidBridge {
		@JavascriptInterface
		fun downloadFile(fileName: String?, mimeType: String?, base64Data: String?) {
			if (base64Data.isNullOrBlank()) {
				return
			}

			val safeName = fileName?.ifBlank { null } ?: "attachment.bin"
			val safeMimeType = mimeType?.ifBlank { null } ?: "application/octet-stream"
			val decoded = try {
				Base64.decode(base64Data, Base64.DEFAULT)
			} catch (_: IllegalArgumentException) {
				runOnUiThread {
					Toast.makeText(this@MainActivity, "Invalid download data.", Toast.LENGTH_LONG).show()
				}
				return
			}

			val savedUri = saveToDownloads(safeName, safeMimeType, decoded)
			if (savedUri != null) {
				runOnUiThread {
					Toast.makeText(
						this@MainActivity,
						"Saved $safeName to Downloads",
						Toast.LENGTH_SHORT,
					).show()
					openDownloadedFile(savedUri, safeMimeType)
				}
				return
			}

			runOnUiThread {
				pendingSaveRequest = PendingSaveRequest(safeName, safeMimeType, decoded)
				createDocumentLauncher.launch(safeName)
			}
		}
	}

	private fun saveToDownloads(fileName: String, mimeType: String, bytes: ByteArray): Uri? {
		if (Build.VERSION.SDK_INT < Build.VERSION_CODES.Q) {
			return null
		}

		return try {
			val values = ContentValues().apply {
				put(MediaStore.MediaColumns.DISPLAY_NAME, fileName)
				put(MediaStore.MediaColumns.MIME_TYPE, mimeType)
				put(MediaStore.MediaColumns.RELATIVE_PATH, "$DOWNLOADS_DIR/$LOOKOUT_SUBDIR")
				put(MediaStore.MediaColumns.IS_PENDING, 1)
			}

			val uri = contentResolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, values) ?: return null
			contentResolver.openOutputStream(uri)?.use { output ->
				output.write(bytes)
			} ?: return null

			val publishValues = ContentValues().apply {
				put(MediaStore.MediaColumns.IS_PENDING, 0)
			}
			contentResolver.update(uri, publishValues, null, null)
			uri
		} catch (_: IOException) {
			null
		} catch (_: SecurityException) {
			null
		}
	}

	private fun openDownloadedFile(uri: Uri, mimeType: String) {
		val openIntent = Intent(Intent.ACTION_VIEW).apply {
			setDataAndType(uri, mimeType)
			addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
		}

		try {
			startActivity(openIntent)
		} catch (_: Exception) {
			Toast.makeText(this, "No app found to open this file.", Toast.LENGTH_SHORT).show()
		}
	}

	private companion object {
		private const val HOST = "localhost"
		private const val BASE_URL = "https://$HOST/index.html?android=1"
		private val DOWNLOADS_DIR = Environment.DIRECTORY_DOWNLOADS
		private const val LOOKOUT_SUBDIR = "LookOut"
	}
}
