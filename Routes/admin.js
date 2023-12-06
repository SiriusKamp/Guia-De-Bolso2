const express = require("express");
const router = express.Router();
const mongoose = require("mongoose");
const ExcelJS = require("exceljs");
const { eAdmin } = require("../helpers/eAdmin");
const { format, parseISO, getYear } = require("date-fns");
const ptBR = require("date-fns/locale/pt-BR");
const Admin = mongoose.model("admins");
const crypto = require("crypto");

router.get("/exportar-admins", eAdmin, async (req, res) => {
  try {
    // Verifique se um ano de nascimento foi fornecido
    const { anoNascimento } = req.query;
    let filtro = {};

    if (anoNascimento) {
      // Converta o ano fornecido para uma data no formato ISO
      const dataInicio = new Date(`${anoNascimento}-01-01T00:00:00.000Z`);
      const dataFim = new Date(`${anoNascimento}-12-31T23:59:59.999Z`);

      // Filtre pelo intervalo de datas
      filtro = {
        nascimento: {
          $gte: dataInicio,
          $lte: dataFim,
        },
      };
    }

    // Busque os admins no banco de dados usando o filtro
    const admins = await Admin.find(filtro).sort({ nascimento: "desc" });

    // Crie um novo workbook Excel
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Admins");

    // Adicione cabeçalhos
    const cabecalhos = ["Email", "Data de Nascimento", "Pontuação"];
    sheet.addRow(cabecalhos);

    // Adicione dados
    admins.forEach((admin) => {
      const linha = [
        admin.email,
        format(new Date(admin.nascimento), "dd/MM/yyyy", { locale: ptBR }),
        admin.pontuacao,
      ];
      sheet.addRow(linha);
    });

    // Crie um stream e envie o arquivo para o cliente
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", "attachment; filename=admins.xlsx");
    await workbook.xlsx.write(res);

    console.log("Exportação concluída com sucesso");
  } catch (err) {
    console.error("Erro ao exportar admins:", err);
    res.status(500).send("Erro ao exportar admins");
  }
});

router.get("/", eAdmin, function (req, res) {
  Admin.find()
    .sort({ nascimento: "desc" })
    .then((admins) => {
      // Mapeia os admins formatando a data de nascimento
      const formattedAdmins = admins.map((admin) => ({
        ...admin._doc,
        nascimento: format(new Date(admin.nascimento), "dd/MM/yyyy", {
          locale: ptBR,
        }),
      }));

      res.render("admin/homeadm", { admins: formattedAdmins });
    })
    .catch((err) => {
      req.flash("error_msg", "Houve um erro ao listar os usuários");
      res.redirect("/admin/");
      console.log(err);
    });
});

router.post("/homeadm/delete", eAdmin, async function (req, res) {
  const adminId = req.body.id;
  console.log(adminId);
  try {
    // Use deleteOne em vez de remove
    const result = await Admin.deleteOne({ _id: adminId });

    if (result.deletedCount === 0) {
      req.flash("error_msg", "Este usuário não existe");
    } else {
      req.flash("success_msg", "Admin deletado com sucesso");
    }

    res.redirect("/admin/");
  } catch (err) {
    console.error(err);
    req.flash("error_msg", "Erro ao excluir admin");
    res.redirect("/admin/");
  }
});

router.post("/remover-nao-confirmados", async (req, res) => {
  try {
    // Remova todos os usuários não confirmados
    await Admin.deleteMany({ isConfirmed: false });

    req.flash("success_msg", "Usuários não confirmados removidos com sucesso");
    res.redirect("/admin/");
  } catch (error) {
    console.error(error);
    req.flash("error_msg", "Houve um erro ao remover usuários não confirmados");
    res.redirect("/admin/");
  }
});

// Rota para transformar um usuário em admin
router.post("/tornar-admin", eAdmin, async (req, res) => {
  const adminId = req.body.id;

  try {
    // Encontrar o admin pelo ID e atualizar as permissões
    const admin = await Admin.findByIdAndUpdate(adminId, { permitions: 1 });

    if (!admin) {
      console.error(err);
      req.flash("error_msg", "Admin não encontrado");
      return res.redirect("/admin/");
    }

    req.flash("success_msg", "Usuário tornou-se admin com sucesso");
    res.redirect("/admin/");
  } catch (err) {
    console.error(err);
    req.flash("error_msg", "Houve um erro ao tornar o usuário um admin");
    res.redirect("/admin/");
  }
});

router.post("/remover-admin", eAdmin, async (req, res) => {
  const adminId = req.body.id;

  try {
    // Encontrar o admin pelo ID e atualizar as permissões
    const admin = await Admin.findByIdAndUpdate(adminId, { permitions: 0 });

    if (!admin) {
      console.error(err);
      req.flash("error_msg", "Admin não encontrado");
      return res.redirect("/admin/");
    }

    req.flash("success_msg", "Permissões removidas com sucesso");
    res.redirect("/admin/");
  } catch (err) {
    console.error(err);
    req.flash("error_msg", "Houve um erro ao remover permissões");
    res.redirect("/admin/");
  }
});

module.exports = router;
