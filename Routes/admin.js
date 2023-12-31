const express = require("express");
const router = express.Router();
const mongoose = require("mongoose");
const ExcelJS = require("exceljs");
const { eAdmin } = require("../helpers/eAdmin");
const { format, parseISO, getYear } = require("date-fns");
const ptBR = require("date-fns/locale/pt-BR");
const Admin = mongoose.model("admins");
const crypto = require("crypto");

router.get("/exportar-admins", async (req, res) => {
  try {
    // Verifique se uma data de criação foi fornecida
    const { datacriacao } = req.query;
    let filtro = {};

    if (datacriacao) {
      // Converta a data fornecida para uma data no formato ISO
      const [dia, mes, ano] = datacriacao.split("/");
      const dataInicio = new Date(ano, mes - 1, dia); // Mês em JavaScript é 0-indexed
      const dataFim = new Date(ano, mes - 1, dia, 23, 59, 59, 999);

      // Filtre pelo intervalo de datas
      filtro = {
        datacriacao: {
          $gte: dataInicio,
          $lte: dataFim,
        },
      };
    }

    // Busque os admins no banco de dados usando o filtro
    const admins = await Admin.find(filtro).sort({ datacriacao: "desc" });

    // Crie um novo workbook Excel
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Admins");

    // Adicione cabeçalhos
    const cabecalhos = ["Data de Nascimento", "Pontuação", "Data de Criação"];
    sheet.addRow(cabecalhos);

    // Adicione dados
    admins.forEach((admin) => {
      const linha = [
        format(new Date(admin.nascimento), "dd/MM/yyyy", { locale: ptBR }),
        admin.pontuacao,
        format(new Date(admin.datacriacao), "dd/MM/yyyy", { locale: ptBR }),
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
    req.flash("error_msg", "coloque a data no formato aa/mm/aaaa");
    res.redirect("/admin/");
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
