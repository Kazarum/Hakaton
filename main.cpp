#include <iostream>
#include "sqlite3.h"
#include <OpenXLSX.hpp>

using namespace OpenXLSX;
using namespace std;


int main() {
	setlocale(LC_ALL, "ru");
	string EXCEL_FILE = "мЮЦПСГЙЮ_МЮ_ЦПСООШоПХЛЕП_БУНДМШУ_ДЮММШУmock.xlsx";

    int rowCounter = 0;

	sqlite3* db;
	sqlite3_open("M:/SQLBases/LearnBase/Learnbase.db", &db);

	string sql_create =
		"CREATE TABLE IF NOT EXISTS load_on_groups ("
		"ID INTEGER PRIMARY KEY AUTOINCREMENT, "
        "яегнм TEXT, "
        "т_на TEXT, "
        "йнд TEXT, "
        "опедлер TEXT, "
        "тюйск TEXT, "
        "яо_рэ TEXT, "
        "яел INTEGER, "
        "цпо INTEGER, "
        "йнк_ярсд INTEGER, "
        "йнк_ба INTEGER, "
        "йнк_хмн INTEGER, "
        "бюп_пюяв INTEGER, "
        "юрр TEXT, "
        "ZET REAL, "
        "кей REAL, "
        "ог REAL, "
        "кп REAL, "
        "юсд_япя REAL, "
        "мю REAL, "
        "HKR REAL, "
        "HPR REAL, "
        "HAT REAL, "
        "H REAL, "
        "йютедпю TEXT);";

        sqlite3_exec(db, sql_create.c_str(), 0, 0, 0);

        XLDocument doc;
		doc.open(EXCEL_FILE);

		auto wks = doc.workbook().worksheet("first");

        sqlite3_exec(db, "BEGIN TRANSACTION;", NULL, NULL, NULL);

        const char* sql_insert = "INSERT INTO load_on_groups (яегнм, т_на, йнд, опедлер, тюйск, яо_рэ, яел, цпо, йнк_ярсд, йнк_ба, йнк_хмн, бюп_пюяв, ATT, ZET, кей, ог, кп, юсд_япя, HA, HKR, HPR, HAT, H, йютедпю) VALUES (-, -, -, -, -, -, -, -, -, -, -, -, -, -, -, -, -, -, -, -, -, -, -, -;";
        sqlite3_stmt* stmt;
        sqlite3_prepare_v2(db, sql_insert, -1, &stmt, NULL);

        for (auto& row : wks.rows()) {
            rowCounter++;
            if (rowCounter <= 2) continue;

            auto cells = row.cells();
            if (cells.empty()) continue;
            sqlite3_bind_text(stmt, 1, cells[0].value().get<string>().c_str(), -1, SQLITE_STATIC); // яегнм
            sqlite3_bind_text(stmt, 2, cells[1].value().get<string>().c_str(), -1, SQLITE_STATIC); // т_на
            sqlite3_bind_text(stmt, 3, cells[2].value().get<string>().c_str(), -1, SQLITE_STATIC); // йнд
            sqlite3_bind_text(stmt, 4, cells[3].value().get<string>().c_str(), -1, SQLITE_STATIC); // опедлер
            sqlite3_bind_text(stmt, 5, cells[4].value().get<string>().c_str(), -1, SQLITE_STATIC); // тюйск
            sqlite3_bind_text(stmt, 6, cells[5].value().get<string>().c_str(), -1, SQLITE_STATIC); // яо_рэ
            sqlite3_bind_int(stmt, 7, cells[6].value().get<int>()); // яел
            sqlite3_bind_int(stmt, 8, cells[7].value().get<int>()); // цпо
            sqlite3_bind_int(stmt, 9, cells[8].value().get<int>()); // йнк_ярсд
            sqlite3_bind_int(stmt, 10, cells[9].value().get<int>()); // йнк_ба
            sqlite3_bind_int(stmt, 11, cells[10].value().get<int>()); // йнк_хмн
            sqlite3_bind_int(stmt, 12, cells[11].value().get<int>()); // бюп_пюяв
            sqlite3_bind_text(stmt, 13, cells[12].value().get<string>().c_str(), -1, SQLITE_STATIC); // ATT
            sqlite3_bind_double(stmt, 14, cells[13].value().get<double>()); // ZET
            sqlite3_bind_double(stmt, 15, cells[14].value().get<double>()); // кей
            sqlite3_bind_double(stmt, 16, cells[15].value().get<double>()); // ог
            sqlite3_bind_double(stmt, 17, cells[16].value().get<double>()); // кп
            sqlite3_bind_double(stmt, 18, cells[17].value().get<double>()); // юсд_япя
            sqlite3_bind_double(stmt, 19, cells[18].value().get<double>()); // HA
            sqlite3_bind_double(stmt, 20, cells[19].value().get<double>()); // HKR
            sqlite3_bind_double(stmt, 21, cells[20].value().get<double>()); // HPR
            sqlite3_bind_double(stmt, 22, cells[21].value().get<double>()); // HAT
            sqlite3_bind_double(stmt, 23, cells[22].value().get<double>()); // H
            sqlite3_bind_text(stmt, 24, cells[23].value().get<string>().c_str(), -1, SQLITE_STATIC); // йютедпю

            sqlite3_step(stmt);
            sqlite3_reset(stmt);
        }

        sqlite3_exec(db, "COMMIT;", NULL, NULL, NULL);
        sqlite3_finalize(stmt);
		sqlite3_close(db);


	return 0;
}