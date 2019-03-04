package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.ClientService;
import com.example.demo.service.FactureService;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.List;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @Autowired
    private FactureService factureService;

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();

        List<Client> clients = clientService.findAllClients();

        writer.println("Id;Nom;Prénom;Date de naissance;Age");
        LocalDate now2 = LocalDate.now();

        for (Client client : clients) {

            writer.println(
                    client.getId() + ";"
                            + client.getNom() + ";"
                            + client.getPrenom() + ";"
                            + client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")) + ";"
                            + (now2.getYear() - client.getDateNaissance().getYear())
            );
        }
    }

    @GetMapping("/clients/xlsx")
    public void clientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        List<Client> clients = clientService.findAllClients();
        LocalDate now = LocalDate.now();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");
        Row headerRow = sheet.createRow(0);


        Cell cellHeaderId = headerRow.createCell(0);
        cellHeaderId.setCellValue("Id");

        Cell cellHeaderPrenom = headerRow.createCell(1);
        cellHeaderPrenom.setCellValue("Prénom");

        Cell cellHeaderNom = headerRow.createCell(2);
        cellHeaderNom.setCellValue("Nom");

        Cell cellHeaderDateNaissance = headerRow.createCell(3);
        cellHeaderDateNaissance.setCellValue("Date de naissance");

        Cell cellHeaderAge = headerRow.createCell(4);
        cellHeaderAge.setCellValue("Age");

        int i = 1;
        for (Client client : clients) {
            Row row = sheet.createRow(i);

            Cell cellId = headerRow.createCell(0);
            cellId.setCellValue(client.getId());

            Cell cellPrenom = row.createCell(1);
            cellPrenom.setCellValue(client.getPrenom());

            Cell cellNom = row.createCell(2);
            cellNom.setCellValue(client.getNom());

            Cell cellDateNaissance = row.createCell(3);
            cellDateNaissance.setCellValue(client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));

            Cell cellAge = row.createCell(4);
            cellAge.setCellValue((now.getYear() - client.getDateNaissance().getYear()));

            i++;


        }

        workbook.write(response.getOutputStream());
        workbook.close();

    }

    @GetMapping("/factures/xlsx")
    public void facturesXlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures.xlsx\"");

        List<Facture> factures = factureService.findAllFactures();
        Workbook workbook = new XSSFWorkbook();


        for (Facture facture : factures) {

            Double prixTotal = 0d;

            Sheet sheet = workbook.createSheet("Factures" + facture.getId());

            Row headerRow = sheet.createRow(0);


            Cell cellHeaderArticle = headerRow.createCell(0);
            cellHeaderArticle.setCellValue("Article");

            Cell cellHeaderQuantite = headerRow.createCell(1);
            cellHeaderQuantite.setCellValue("Quantité");

            Cell cellHeaderPrixUnitaire = headerRow.createCell(2);
            cellHeaderPrixUnitaire.setCellValue("Prix Unitaire");

            Cell cellHeaderPrixLigne = headerRow.createCell(3);
            cellHeaderPrixLigne.setCellValue("Prix Ligne");

            int i = 2;

            for (LigneFacture ligneFacture : facture.getLigneFactures()) {
                Row row = sheet.createRow(i);

                Cell cellArticle = row.createCell(0);
              cellArticle.setCellValue(ligneFacture.getArticle().getLibelle());



                Cell cellQuantite = row.createCell(1);
                cellQuantite.setCellValue(ligneFacture.getQuantite());

                Cell cellPrixUnitaire = row.createCell(2);
                cellPrixUnitaire.setCellValue(ligneFacture.getArticle().getPrix());

                Cell cellPrixLigne = row.createCell(3);
                cellPrixLigne.setCellValue(ligneFacture.getArticle().getPrix() * ligneFacture.getQuantite());

                prixTotal += ligneFacture.getArticle().getPrix() * ligneFacture.getQuantite();



                i++;




            }
            Row row = sheet.createRow(i+2);
            Cell cellPrixTotal = row.createCell (3);
            cellPrixTotal.setCellValue(prixTotal);

            workbook.write(response.getOutputStream());
            workbook.close();
        }




    }
}
