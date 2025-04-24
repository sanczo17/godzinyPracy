package com.example.godzinypracy

import android.app.TimePickerDialog
import android.content.Intent
import android.os.Bundle
import android.view.View
import android.widget.AdapterView
import android.widget.ArrayAdapter
import android.widget.Button
import android.widget.EditText
import android.widget.Spinner
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import com.google.android.gms.auth.api.signin.GoogleSignIn
import com.google.android.gms.auth.api.signin.GoogleSignInAccount
import com.google.android.gms.auth.api.signin.GoogleSignInClient
import com.google.android.gms.auth.api.signin.GoogleSignInOptions
import com.google.android.gms.common.SignInButton
import com.google.android.gms.common.api.Scope
import com.google.api.client.http.javanet.NetHttpTransport
import com.google.api.client.json.gson.GsonFactory
import com.google.api.client.googleapis.extensions.android.gms.auth.GoogleAccountCredential
import com.google.api.services.sheets.v4.Sheets
import com.google.api.services.sheets.v4.SheetsScopes
import com.google.api.services.sheets.v4.model.AddSheetRequest
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest
import com.google.api.services.sheets.v4.model.CellData
import com.google.api.services.sheets.v4.model.CellFormat
import com.google.api.services.sheets.v4.model.ExtendedValue
import com.google.api.services.sheets.v4.model.GridData
import com.google.api.services.sheets.v4.model.GridProperties
import com.google.api.services.sheets.v4.model.Request
import com.google.api.services.sheets.v4.model.RowData
import com.google.api.services.sheets.v4.model.Sheet
import com.google.api.services.sheets.v4.model.SheetProperties
import com.google.api.services.sheets.v4.model.Spreadsheet
import com.google.api.services.sheets.v4.model.UpdateCellsRequest
import com.google.api.services.sheets.v4.model.ValueRange
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.text.SimpleDateFormat
import java.util.Calendar
import java.util.Collections
import java.util.Date
import java.util.Locale
import android.widget.*
import com.google.android.gms.common.api.ApiException
import java.util.*

class MainActivity : AppCompatActivity() {

    companion object {
        private const val RC_SIGN_IN = 9001
        private const val SPREADSHEET_ID = "1Q2vFanQRkU9c9dJTfAlwXlU4wT5w94_LsECqv6ME-ok" // ID arkusza z URL
    }

    private lateinit var dateTextView: TextView
    private lateinit var startTimeEditText: EditText
    private lateinit var endTimeEditText: EditText
    private lateinit var descriptionEditText: EditText
    private lateinit var counterEditText: EditText
    private lateinit var saveButton: Button
    private lateinit var dayOffButton: Button
    private lateinit var signInButton: SignInButton
    private lateinit var operatorTextView: TextView
    private lateinit var currentSheetNameTextView: TextView
    private lateinit var statusSpinner: Spinner

    private lateinit var mGoogleSignInClient: GoogleSignInClient
    private var sheetsService: Sheets? = null

    // Przechowuje aktualną nazwę arkusza (miesiąc/rok)
    private var currentSheetName: String = ""

    // Operator
    private var currentOperator: String = "Mateusz Dylewski" // Domyślny operator

    // Status dnia (normalny, wolne, urlop, L4)
    private var dayStatus: String = ""
    private val statusOptions = listOf("Normalny dzień", "Wolne", "Urlop", "L4", "Święto")

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        // Inicjalizacja elementów UI
        dateTextView = findViewById(R.id.dateTextView)
        startTimeEditText = findViewById(R.id.startTimeEditText)
        endTimeEditText = findViewById(R.id.endTimeEditText)
        descriptionEditText = findViewById(R.id.descriptionEditText)
        counterEditText = findViewById(R.id.counterEditText)
        saveButton = findViewById(R.id.saveButton)
        dayOffButton = findViewById(R.id.dayOffButton)
        signInButton = findViewById(R.id.signInButton)
        operatorTextView = findViewById(R.id.operatorTextView)
        currentSheetNameTextView = findViewById(R.id.currentSheetNameTextView)
        statusSpinner = findViewById(R.id.statusSpinner)

        // Inicjalizacja spinnera statusu
        val adapter = ArrayAdapter(this, android.R.layout.simple_spinner_item, statusOptions)
        adapter.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item)
        statusSpinner.adapter = adapter

        // Obsługa wyboru statusu
        statusSpinner.onItemSelectedListener = object : AdapterView.OnItemSelectedListener {
            override fun onItemSelected(parent: AdapterView<*>?, view: View?, position: Int, id: Long) {
                dayStatus = statusOptions[position]

                // Aktualizuj interfejs w zależności od wybranego statusu
                when (dayStatus) {
                    "Normalny dzień" -> {
                        startTimeEditText.isEnabled = true
                        endTimeEditText.isEnabled = true
                        counterEditText.isEnabled = true
                        descriptionEditText.isEnabled = true
                    }
                    else -> {
                        // Dla dni wolnych, urlopu, L4, i świąt
                        startTimeEditText.isEnabled = false
                        endTimeEditText.isEnabled = false
                        counterEditText.isEnabled = false
                        descriptionEditText.setText("")

                        // Ustaw domyślny opis w zależności od statusu
                        when (dayStatus) {
                            "Wolne" -> descriptionEditText.setText("")
                            "Urlop" -> descriptionEditText.setText("Urlop")
                            "L4" -> descriptionEditText.setText("L4")
                            "Święto" -> descriptionEditText.setText("Święto")
                        }
                    }
                }
            }

            override fun onNothingSelected(parent: AdapterView<*>?) {
                // Nic nie rób
            }
        }

        // Ustaw dzisiejszą datę
        val dateFormat = SimpleDateFormat("dd-MM-yyyy", Locale.getDefault())
        val currentDate = dateFormat.format(Date())
        dateTextView.text = "Data: $currentDate"

        // Ustawienie nazwy bieżącego arkusza (miesiąc/rok)
        updateCurrentSheetName()

        // Konfiguracja Google Sign-In
        val gso = GoogleSignInOptions.Builder(GoogleSignInOptions.DEFAULT_SIGN_IN)
            .requestEmail()
            .requestScopes(Scope(SheetsScopes.SPREADSHEETS))
            .requestIdToken("918239658749-b03ss40uv6k0q6oa414g2snn58aarfph.apps.googleusercontent.com")
            .build()
        mGoogleSignInClient = GoogleSignIn.getClient(this, gso)

        // Przypisanie akcji do przycisków
        signInButton.setOnClickListener { signIn() }

        startTimeEditText.setOnClickListener { showTimePickerDialog(startTimeEditText) }
        endTimeEditText.setOnClickListener { showTimePickerDialog(endTimeEditText) }

        saveButton.setOnClickListener { saveWorkTime() }
        dayOffButton.setOnClickListener {
            // Ustaw spinner na "Wolne"
            val wolneIndex = statusOptions.indexOf("Wolne")
            if (wolneIndex >= 0) {
                statusSpinner.setSelection(wolneIndex)
            }
            // Zapisz dzień wolny
            saveWorkTime()
        }

        // Aktualizuj operatora
        operatorTextView.text = "Operator: $currentOperator"

        // Sprawdź, czy użytkownik jest już zalogowany
        val account = GoogleSignIn.getLastSignedInAccount(this)
        if (account != null) {
            updateUI(account)
        }
    }

    // Aktualizacja nazwy bieżącego arkusza
    private fun updateCurrentSheetName() {
        val calendar = Calendar.getInstance()
        val monthYear = getMonthYearString(calendar)
        currentSheetName = monthYear
        currentSheetNameTextView.text = "Arkusz: $monthYear"
    }

    // Uzyskanie nazwy miesiąca i roku w formacie "MiesiącRRRR"
    private fun getMonthYearString(calendar: Calendar): String {
        val locale = Locale("pl", "PL")
        val monthFormat = SimpleDateFormat("MMMM", locale)
        val yearFormat = SimpleDateFormat("yyyy", locale)

        val month = monthFormat.format(calendar.time)
        val year = yearFormat.format(calendar.time)

        // Pierwszy znak miesiąca jako wielka litera, reszta małe
        val capitalizedMonth = month.replaceFirstChar {
            if (it.isLowerCase()) it.titlecase(locale) else it.toString()
        }

        return "$capitalizedMonth$year"
    }

    // Metoda do wyświetlania selektora czasu
    private fun showTimePickerDialog(editText: EditText) {
        val calendar = Calendar.getInstance()
        val hour = calendar.get(Calendar.HOUR_OF_DAY)
        val minute = calendar.get(Calendar.MINUTE)

        TimePickerDialog(this,
            { _, hourOfDay, minute ->
                val time = String.format(Locale.getDefault(), "%02d:%02d", hourOfDay, minute)
                editText.setText(time)
            }, hour, minute, true).show()
    }

    // Logowanie użytkownika
    private fun signIn() {
        val signInIntent = mGoogleSignInClient.signInIntent
        startActivityForResult(signInIntent, RC_SIGN_IN)
    }

    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)

        if (requestCode == RC_SIGN_IN) {
            val task = GoogleSignIn.getSignedInAccountFromIntent(data)
            try {
                val account = task.getResult(ApiException::class.java)
                updateUI(account)
            } catch (e: Exception) {
                updateUI(null)
                Toast.makeText(this, "Błąd logowania: ${e.message}", Toast.LENGTH_SHORT).show()
            }
        }
    }

    // Aktualizacja UI po zalogowaniu
    private fun updateUI(account: GoogleSignInAccount?) {
        if (account != null) {
            signInButton.visibility = View.GONE
            saveButton.isEnabled = true
            dayOffButton.isEnabled = true

            // Aktualizuj operatora jeśli chcesz używać nazwy zalogowanego użytkownika
            // currentOperator = account.displayName ?: "Nieznany operator"
            // operatorTextView.text = "Operator: $currentOperator"

            // Inicjalizacja Sheets API
            initializeSheetsApi(account)

            // Sprawdź i utwórz arkusz dla bieżącego miesiąca jeśli nie istnieje
            lifecycleScope.launch {
                try {
                    val sheetExists = checkIfSheetExists(currentSheetName)
                    if (!sheetExists) {
                        createNewMonthSheet(currentSheetName)
                    }

                    // Załaduj dane z dzisiejszego dnia
                    loadTodayData()
                } catch (e: Exception) {
                    withContext(Dispatchers.Main) {
                        Toast.makeText(this@MainActivity, "Błąd: ${e.message}", Toast.LENGTH_SHORT).show()
                    }
                }
            }
        } else {
            signInButton.visibility = View.VISIBLE
            saveButton.isEnabled = false
            dayOffButton.isEnabled = false
            startTimeEditText.isEnabled = false
            endTimeEditText.isEnabled = false
            descriptionEditText.isEnabled = false
            counterEditText.isEnabled = false
            statusSpinner.isEnabled = false
        }
    }

    // Inicjalizacja Sheets API
    private fun initializeSheetsApi(account: GoogleSignInAccount) {
        val credential = GoogleAccountCredential.usingOAuth2(
            this, Collections.singleton(SheetsScopes.SPREADSHEETS))
        credential.selectedAccount = account.account

        sheetsService = Sheets.Builder(
            NetHttpTransport(),
            GsonFactory(),
            credential)
            .setApplicationName("Godziny Pracy")
            .build()
    }

    // Sprawdzenie czy arkusz o podanej nazwie istnieje
    private suspend fun checkIfSheetExists(sheetName: String): Boolean = withContext(Dispatchers.IO) {
        try {
            val response = sheetsService?.spreadsheets()?.get(SPREADSHEET_ID)?.execute()
            val sheets = response?.sheets ?: emptyList()

            return@withContext sheets.any { it.properties?.title == sheetName }
        } catch (e: Exception) {
            throw Exception("Błąd sprawdzania arkusza: ${e.message}")
        }
    }

    // Tworzenie nowego arkusza dla miesiąca
    private suspend fun createNewMonthSheet(sheetName: String) = withContext(Dispatchers.IO) {
        try {
            // Dodaj nowy arkusz
            val addSheetRequest = AddSheetRequest().apply {
                properties = SheetProperties().apply {
                    title = sheetName
                    gridProperties = GridProperties().apply {
                        rowCount = 34 // Nagłówek + 31 dni + suma + wiersz zapasowy
                        columnCount = 7 // A-G
                    }
                }
            }

            val batchUpdateRequest = BatchUpdateSpreadsheetRequest().apply {
                requests = listOf(Request().setAddSheet(addSheetRequest))
            }

            sheetsService?.spreadsheets()?.batchUpdate(SPREADSHEET_ID, batchUpdateRequest)?.execute()

            // Dodaj nagłówki i formuły
            setupNewSheetHeaders(sheetName)
            setupNewSheetFormulas(sheetName)

            withContext(Dispatchers.Main) {
                Toast.makeText(this@MainActivity, "Utworzono nowy arkusz: $sheetName", Toast.LENGTH_SHORT).show()
            }
        } catch (e: Exception) {
            throw Exception("Błąd tworzenia arkusza: ${e.message}")
        }
    }

    // Konfiguracja nagłówków w nowym arkuszu
    private suspend fun setupNewSheetHeaders(sheetName: String) = withContext(Dispatchers.IO) {
        val headers = listOf(
            "Data", "Operator", "Rozp. pracy", "Zakoń. pracy", "Stan Licznika", "Czynności", "Podsumowanie"
        )

        val headerRange = "$sheetName!A1:G1"
        val headerValues = ValueRange().setValues(listOf(headers))

        sheetsService?.spreadsheets()?.values()
            ?.update(SPREADSHEET_ID, headerRange, headerValues)
            ?.setValueInputOption("RAW")
            ?.execute()

        // Dodaj daty do kolumny A (od A2 do A32)
        addDatesToDaysColumn(sheetName)
    }

    // Dodawanie dat do kolumny dni
    private suspend fun addDatesToDaysColumn(sheetName: String) = withContext(Dispatchers.IO) {
        val calendar = Calendar.getInstance()
        val year = calendar.get(Calendar.YEAR)
        val month = calendar.get(Calendar.MONTH)

        // Ustaw kalendarz na pierwszy dzień miesiąca
        calendar.set(year, month, 1)

        // Pobierz liczbę dni w miesiącu
        val daysInMonth = calendar.getActualMaximum(Calendar.DAY_OF_MONTH)

        val dateFormat = SimpleDateFormat("dd-MM-yyyy", Locale.getDefault())
        val dates = ArrayList<List<Any>>()

        for (day in 1..daysInMonth) {
            calendar.set(Calendar.DAY_OF_MONTH, day)
            dates.add(listOf(dateFormat.format(calendar.time)))
        }

        val datesRange = "$sheetName!A2:A${daysInMonth + 1}"
        val datesValues = ValueRange().setValues(dates)

        sheetsService?.spreadsheets()?.values()
            ?.update(SPREADSHEET_ID, datesRange, datesValues)
            ?.setValueInputOption("RAW")
            ?.execute()
    }

    // Konfiguracja formuł w nowym arkuszu
    private suspend fun setupNewSheetFormulas(sheetName: String) = withContext(Dispatchers.IO) {
        // Znajdź ID arkusza
        val response = sheetsService?.spreadsheets()?.get(SPREADSHEET_ID)?.execute()
        val sheet = response?.sheets?.find { it.properties?.title == sheetName }
        val sheetId = sheet?.properties?.sheetId ?: return@withContext

        // Formuła dla komórki G2 (i pozostałych w kolumnie G)
        val formulaG2 = "=IF(OR(B2=\"Wolne\"; B2=\"Święto\";B2=\"L4\"; C2=\"\"; D2=\"\"); \"\"; (D2-C2)*24)"

        // Formuła dla komórki G34 (suma godzin)
        val sumFormula = "=SUM(G2:G33)"
        val sumLabel = "Łącznie czas pracy"

        // Dodaj formułę dla każdej komórki G2:G33
        for (row in 2..33) {
            val range = "$sheetName!G$row"
            val formula = if (row == 33) sumFormula else formulaG2.replace("2", row.toString())

            val formulaValues = ValueRange().setValues(listOf(listOf(formula)))

            sheetsService?.spreadsheets()?.values()
                ?.update(SPREADSHEET_ID, range, formulaValues)
                ?.setValueInputOption("USER_ENTERED") // Ważne aby użyć USER_ENTERED dla formuł
                ?.execute()
        }

        // Dodaj etykietę w F33
        val labelRange = "$sheetName!F33"
        val labelValues = ValueRange().setValues(listOf(listOf(sumLabel)))

        sheetsService?.spreadsheets()?.values()
            ?.update(SPREADSHEET_ID, labelRange, labelValues)
            ?.setValueInputOption("RAW")
            ?.execute()
    }

    // Ładowanie danych z dzisiejszego dnia
    private suspend fun loadTodayData() = withContext(Dispatchers.IO) {
        try {
            val dateFormat = SimpleDateFormat("dd-MM-yyyy", Locale.getDefault())
            val currentDate = dateFormat.format(Date())

            val response = sheetsService?.spreadsheets()?.values()
                ?.get(SPREADSHEET_ID, "$currentSheetName!A:G")
                ?.execute()

            val values = response?.getValues()

            if (values != null) {
                // Szukaj wiersza z dzisiejszą datą
                for (i in values.indices) {
                    val row = values[i]
                    if (row.size > 0 && row[0] == currentDate) {
                        val rowData = row

                        withContext(Dispatchers.Main) {
                            // Ustaw status dnia
                            if (rowData.size > 1) {
                                val status = when {
                                    rowData[1] == "Wolne" -> "Wolne"
                                    rowData[1] == "Urlop" -> "Urlop"
                                    rowData[1] == "L4" -> "L4"
                                    rowData.size > 5 && rowData[5] == "Święto" -> "Święto"
                                    else -> "Normalny dzień"
                                }

                                val statusIndex = statusOptions.indexOf(status)
                                if (statusIndex >= 0) {
                                    statusSpinner.setSelection(statusIndex)
                                }
                            }

                            // Uzupełnij pola danymi z arkusza
                            if (rowData.size > 2 && rowData[2].toString().isNotEmpty())
                                startTimeEditText.setText(rowData[2].toString())

                            if (rowData.size > 3 && rowData[3].toString().isNotEmpty())
                                endTimeEditText.setText(rowData[3].toString())

                            if (rowData.size > 4 && rowData[4].toString().isNotEmpty())
                                counterEditText.setText(rowData[4].toString())

                            if (rowData.size > 5 && rowData[5].toString().isNotEmpty())
                                descriptionEditText.setText(rowData[5].toString())
                        }

                        break
                    }
                }
            }
        } catch (e: Exception) {
            withContext(Dispatchers.Main) {
                Toast.makeText(this@MainActivity, "Błąd ładowania danych: ${e.message}", Toast.LENGTH_SHORT).show()
            }
        }
    }

    // Zapisywanie czasu pracy
    private fun saveWorkTime() {
        lifecycleScope.launch {
            try {
                val dateFormat = SimpleDateFormat("dd-MM-yyyy", Locale.getDefault())
                val currentDate = dateFormat.format(Date())

                // Przygotuj dane do zapisu
                val rowData = mutableListOf<Any>(
                    currentDate,
                    when (dayStatus) {
                        "Wolne", "Urlop", "L4", "Święto" -> dayStatus
                        else -> currentOperator
                    }
                )

                // Dodaj czas rozpoczęcia
                if (dayStatus == "Normalny dzień") {
                    if (startTimeEditText.text.toString().isEmpty()) {
                        withContext(Dispatchers.Main) {
                            Toast.makeText(this@MainActivity, "Wprowadź czas rozpoczęcia pracy", Toast.LENGTH_SHORT).show()
                        }
                        return@launch
                    }
                    rowData.add(startTimeEditText.text.toString())
                } else {
                    rowData.add("")
                }

                // Dodaj czas zakończenia
                if (dayStatus == "Normalny dzień") {
                    if (endTimeEditText.text.toString().isEmpty()) {
                        withContext(Dispatchers.Main) {
                            Toast.makeText(this@MainActivity, "Wprowadź czas zakończenia pracy", Toast.LENGTH_SHORT).show()
                        }
                        return@launch
                    }
                    rowData.add(endTimeEditText.text.toString())
                } else {
                    rowData.add("")
                }

                // Dodaj stan licznika
                rowData.add(counterEditText.text.toString())

                // Dodaj opis czynności
                rowData.add(descriptionEditText.text.toString())

                // Kolumna G (podsumowanie) jest obliczana przez formułę, więc jej nie ustawiamy

                // Znajdź wiersz dla dzisiejszej daty lub użyj odpowiedniego wiersza na podstawie dnia miesiąca
                val day = currentDate.substring(0, 2).toInt()
                val rowIndex = day + 1 // +1 bo wiersz 1 to nagłówek

                // Zaktualizuj dane w arkuszu
                val range = "$currentSheetName!A$rowIndex:F$rowIndex"
                val valueRange = ValueRange().setValues(listOf(rowData))

                sheetsService?.spreadsheets()?.values()
                    ?.update(SPREADSHEET_ID, range, valueRange)
                    ?.setValueInputOption("RAW")
                    ?.execute()

                withContext(Dispatchers.Main) {
                    Toast.makeText(this@MainActivity, "Dane zapisane pomyślnie", Toast.LENGTH_SHORT).show()
                }
            } catch (e: Exception) {
                withContext(Dispatchers.Main) {
                    Toast.makeText(this@MainActivity, "Błąd zapisu: ${e.message}", Toast.LENGTH_SHORT).show()
                }
            }
        }
    }
}