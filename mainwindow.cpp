#include <QFileDialog>
#include <QTextStream>
#include <QDateTime>
#include <QMessageBox>
#include <QVBoxLayout>
#include "mainwindow.h"
#include "./ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    interval = 0;

    ui->setupUi(this);

    // Zablokowanie rozmiaru na obecne wymiary okna
    setFixedSize(size());

    ui->generateExcelFileButton->setEnabled(false);

    //connect(ui->hourSpinBox, &QSpinBox::valueChanged, this, &MainWindow::updateSeconds);
    //connect(ui->minuteSpinBox, &QSpinBox::valueChanged, this, &MainWindow::updateSeconds);
    //connect(ui->secondSpinBox, &QSpinBox::valueChanged, this, &MainWindow::updateSeconds);

}

MainWindow::~MainWindow()
{
    delete ui;
    if (file.isOpen())
    {
        file.close();
    }
}


void MainWindow::on_openFileButton_clicked()
{
    if (file.isOpen())
    {
        file.close();
    }

    fileName = QFileDialog::getOpenFileName(this,
                                                    "Otwórz plik",
                                                    QString(),
                                                    "Pliki tekstowe (*.txt)");
    if (!fileName.isEmpty())
    {
        // Plik został wybrany, kontynuuj z obsługą pliku
        file.setFileName(fileName);

        if (file.open(QIODevice::ReadOnly | QIODevice::Text))
        {
            ui->generateExcelFileButton->setEnabled(true);
            // Plik został pomyślnie otwarty, możesz teraz czytać z pliku
        }
        else
        {
            // Informacja, że pliku nie udało się otworzyć
            QMessageBox::warning(this, "Ostrzeżenie", "Nie można otworzyć pliku.");
        }
    }
    else
    {
        // Użytkownik anulował wybór pliku
        ui->generateExcelFileButton->setEnabled(false);
    }
}


void MainWindow::on_generateExcelFileButton_clicked()
{
    QXlsx::Document xlsx;

    // Formatowanie komórek
    QXlsx::Format centerFormat;
    centerFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    centerFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);

    // Formatowanie nagłówków z pogrubioną czcionką, wyśrodkowaniem i zielonym tłem
    QXlsx::Format headerFormat = centerFormat;
    headerFormat.setFontBold(true);
    headerFormat.setPatternBackgroundColor(QColor(146, 208, 80)); // Zielone tło
    headerFormat.setBorderStyle(QXlsx::Format::BorderThin); // Ustawienie grubego obramowania

    // Ustawienie domyślnego koloru czcionki na czarny
    headerFormat.setFontColor(QColor(0, 0, 0));

    // Zapis nagłówków z formatowaniem
    xlsx.write("A1", "Data", headerFormat);
    xlsx.write("B1", "Godzina", headerFormat);
    xlsx.write("C1", "Komunikat", headerFormat);
    xlsx.write("D1", "Opóźnienie", headerFormat);

    file.seek(0); // Przewijanie pliku na początek
    QTextStream in(&file);
    QString currentLine;
    QDateTime currentTime;
    int row = 2;

    const int maxEventDifferenceSec = ui->maxEventDifferenceSecSpinBox->value(); // Maksymalna różnica czasu
    const int minEvents = ui->minEventsSpinBox->value(); // Minimalna liczba zdarzeń
    QList<QString> eventSeries;
    QDateTime lastEventTime;
    int eventCount = 0;

    QXlsx::Format blueBackgroundFormat = centerFormat;
    blueBackgroundFormat.setPatternBackgroundColor(QColor(141, 180, 226)); // Jasnoniebieskie tło
    blueBackgroundFormat.setBorderStyle(QXlsx::Format::BorderThin); // Ustawienie cieńszego obramowania

    QXlsx::Format grayBackgroundFormat = centerFormat;
    grayBackgroundFormat.setPatternBackgroundColor(QColor(191, 191, 191)); // Jasnoszare tło
    grayBackgroundFormat.setBorderStyle(QXlsx::Format::BorderThin); // Ustawienie cieńszego obramowania

    bool useBlueBackground = true;

    while (!in.atEnd())
    {
        currentLine = in.readLine();
        currentTime = QDateTime::fromString(currentLine.left(19), "yyyy-MM-dd HH:mm:ss");

        int secondsDiff = lastEventTime.secsTo(currentTime);
        if (!lastEventTime.isValid() || (secondsDiff <= maxEventDifferenceSec || (secondsDiff < 0 && -secondsDiff < 86400 - maxEventDifferenceSec)))
        {
            eventSeries.append(currentLine);
            eventCount++;
        }
        else
        {
            if (eventCount >= minEvents) {
                // Zmień tło komórek tylko na początku nowej serii
                QXlsx::Format currentBackgroundFormat = useBlueBackground ? blueBackgroundFormat : grayBackgroundFormat;

                for (const QString &eventLine : eventSeries) {
                    processAndWriteLine(eventLine, QDateTime::fromString(eventLine.left(19), "yyyy-MM-dd HH:mm:ss"), row, xlsx, currentBackgroundFormat);
                    row++;
                }
                useBlueBackground = !useBlueBackground;
            }
            eventSeries.clear();
            eventCount = 1;
            eventSeries.append(currentLine);
        }
        lastEventTime = currentTime;
    }

    if (eventCount >= minEvents) {
        // Zmień tło komórek tylko na początku nowej serii
        QXlsx::Format currentBackgroundFormat = useBlueBackground ? blueBackgroundFormat : grayBackgroundFormat;

        for (const QString &eventLine : eventSeries) {
            processAndWriteLine(eventLine, QDateTime::fromString(eventLine.left(19), "yyyy-MM-dd HH:mm:ss"), row, xlsx, currentBackgroundFormat);
            row++;
        }
    }

    xlsx.autosizeColumnWidth();

    QString xlsxFileName = "ping_stats.xlsx";
    if (xlsx.saveAs(xlsxFileName))
    {
        qDebug() << "Plik XLSX został zapisany jako" << xlsxFileName;
    }
    else
    {
        qDebug() << "Błąd zapisu pliku XLSX";
    }
}

void MainWindow::processAndWriteLine(const QString &line, const QDateTime &time, int row, QXlsx::Document &xlsx, const QXlsx::Format &fontFormat)
{
    QStringList parts = line.split(" - ");
    QString description = parts.size() >= 2 ? parts[1] : "";

    // Obsługa komunikatu zawierającego "ms"
    if (description.contains("ms"))
    {
        QStringList descriptionParts = description.split(":");
        if (descriptionParts.size() == 2)
        {
            QString desc = descriptionParts[0].trimmed();
            QString timeInMs = descriptionParts[1].trimmed();

            xlsx.write(row, 1, time.toString("yyyy-MM-dd"), fontFormat);
            xlsx.write(row, 2, time.toString("HH:mm:ss"), fontFormat);
            xlsx.write(row, 3, desc, fontFormat);
            xlsx.write(row, 4, timeInMs, fontFormat);
        }
    }
    else
    {
        xlsx.write(row, 1, time.toString("yyyy-MM-dd"), fontFormat);
        xlsx.write(row, 2, time.toString("HH:mm:ss"), fontFormat);
        xlsx.write(row, 3, description, fontFormat);
        xlsx.write(row, 4, "", fontFormat); // Kolumna D pozostaje pusta
    }
}

/*void MainWindow::updateSeconds()
{
    int hours = ui->hourSpinBox->value();
    int minutes = ui->minuteSpinBox->value();
    int seconds = ui->secondSpinBox->value();

    interval = hours * 3600 + minutes * 60 + seconds;
}*/


