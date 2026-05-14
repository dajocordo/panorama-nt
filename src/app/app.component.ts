import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { MessageService } from 'primeng/api';
import { ButtonModule } from 'primeng/button';
import { ToastModule } from 'primeng/toast';

// Importaciones para el documento
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak } from 'docx';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-root',
  standalone: true, // Importante en Angular 19
  imports: [
    CommonModule,
    FormsModule,
    ButtonModule,
    ToastModule
  ],
  providers: [MessageService],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent implements OnInit {
  title = 'panorama-nt';
  buscarLibro: string = '';
  expandirTodo: boolean = false;
  libros: any = []
  mostrarAlert: boolean = false;

  alertConfig = { severity: '', summary: '', detail: '' };

  copiarInfo(libro: any) {
    const info = `${libro.libro}

Autor: ${libro.autor}
Fecha: ${libro.fecha}
Destinatarios: ${libro.destinatarios}
Propósito: ${libro.proposito}
Tema: ${libro.tema}

Bosquejo:
${libro.bosquejo.map((punto: string) => `• ${punto}`).join('\n')}`;

    // (Mantén tu lógica de construcción de 'info')

    navigator.clipboard.writeText(info)
      .then(() => {
        this.lanzarNotificacion('success', '¡Copiado!', 'Información al portapapeles');
      })
      .catch(() => {
        this.lanzarNotificacion('error', 'Error', 'No se pudo copiar');
      });
  }

  lanzarNotificacion(severity: string, summary: string, detail: string) {
    this.alertConfig = { severity, summary, detail };
    this.mostrarAlert = true;

    // En móviles 2.5 o 3 segundos es el tiempo perfecto
    setTimeout(() => {
      this.mostrarAlert = false;
    }, 2500);
  }
  ngOnInit() {
    this.llenarLibros();
  }

  descargarTodosLosLibros() {
    const secciones: any[] = [];

    // Iteramos sobre todos los libros que tienes en memoria
    this.libros.forEach((libro: any, index: number) => {

      // 1. Título del Libro
      secciones.push(
        new Paragraph({
          text: libro.libro.toUpperCase(),
          heading: HeadingLevel.HEADING_1,
          alignment: AlignmentType.CENTER,
          spacing: { before: index === 0 ? 0 : 400, after: 400 },
        })
      );

      // 2. Información General
      const campos = [
        { k: 'Autor', v: libro.autor },
        { k: 'Fecha', v: libro.fecha },
        { k: 'Destinatarios', v: libro.destinatarios },
        { k: 'Propósito', v: libro.proposito },
        { k: 'Tema', v: libro.tema }
      ];

      campos.forEach(campo => {
        secciones.push(new Paragraph({
          children: [
            new TextRun({ text: `${campo.k}: `, bold: true }),
            new TextRun({ text: campo.v || 'N/A' }),
          ],
          spacing: { after: 120 },
        }));
      });

      // 3. Bosquejo
      secciones.push(new Paragraph({
        text: "Bosquejo",
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 400, after: 200 },
      }));

      libro.bosquejo.forEach((punto: string) => {
        secciones.push(new Paragraph({
          text: punto,
          bullet: { level: 0 },
        }));
      });

      // 4. Salto de página (excepto después del último libro)
      if (index < this.libros.length - 1) {
        secciones.push(new Paragraph({
          children: [new PageBreak()],
        }));
      }
    });

    // Crear el documento final
    const doc = new Document({
      sections: [{
        properties: {},
        children: secciones,
      }],
    });

    // Guardar archivo
    Packer.toBlob(doc).then((blob) => {
      saveAs(blob, `Panorama_Completo.docx`);
      this.lanzarNotificacion('success', '¡Listo!', `Se exportaron ${this.libros.length} al portapapeles`);
    });
  }
  // Helper para crear líneas "Etiqueta: Valor" con estilo
  private crearLineaInfo(etiqueta: string, valor: string) {
    return [
      new Paragraph({
        children: [
          new TextRun({ text: `${etiqueta}: `, bold: true }),
          new TextRun({ text: valor || 'N/A' }),
        ],
        spacing: { after: 120 },
      })
    ];
  }



  get librosFiltrados() {
    return this.libros.filter((libro: any) =>
      libro.libro.toLowerCase().includes(this.buscarLibro.toLowerCase()) ||
      libro.autor.toLowerCase().includes(this.buscarLibro.toLowerCase())
    );
  }

  toggleExpandirTodo() {
    this.expandirTodo = !this.expandirTodo;
  }



  llenarLibros() {
    this.libros = [
      {
        libro: 'Evangelio según Mateo',
        autor: 'Mateo (Leví)',
        fecha: '60–65 d.C.',
        destinatarios: 'Principalmente judíos',
        proposito: 'Presentar a Jesús como el Mesías prometido y Rey de Israel',
        tema: 'Jesucristo, el Rey prometido',
        bosquejo: [
          'Nacimiento y preparación del Rey (1–4)',
          'Ministerio del Rey en Galilea (5–18)',
          'Rechazo del Rey (19–25)',
          'Muerte y resurrección del Rey (26–28)'
        ]
      },
      {
        libro: 'Evangelio según Marcos',
        autor: 'Juan Marcos',
        fecha: '55–65 d.C.',
        destinatarios: 'Romanos y gentiles',
        proposito: 'Mostrar a Jesús como el Siervo poderoso de Dios',
        tema: 'Jesucristo, el Siervo obediente',
        bosquejo: [
          'Preparación del Siervo (1)',
          'Ministerio del Siervo (2–10)',
          'Sacrificio del Siervo (11–16)'
        ]
      },
      {
        libro: 'Evangelio según Lucas',
        autor: 'Lucas, médico',
        fecha: '60–63 d.C.',
        destinatarios: 'Teófilo y creyentes gentiles',
        proposito: 'Presentar un relato ordenado y completo de la vida de Jesús',
        tema: 'Jesucristo, el Hijo del Hombre',
        bosquejo: [
          'Nacimiento y preparación de Jesús (1–4)',
          'Ministerio en Galilea (5–9)',
          'Camino a Jerusalén (10–19)',
          'Pasión y resurrección (20–24)'
        ]
      },
      {
        libro: 'Evangelio según Juan',
        autor: 'Juan, el apóstol',
        fecha: '85–95 d.C.',
        destinatarios: 'Creyentes y no creyentes en general',
        proposito: 'Demostrar que Jesús es el Hijo de Dios y llevar a la fe',
        tema: 'Jesucristo, el Hijo de Dios',
        bosquejo: [
          'Prólogo: el Verbo eterno (1)',
          'Ministerio público de Jesús (2–12)',
          'Enseñanzas a los discípulos (13–17)',
          'Muerte y resurrección (18–21)'
        ]
      },
      {
        libro: 'Hechos de los Apóstoles',
        autor: 'Lucas',
        fecha: '62–64 d.C.',
        destinatarios: 'Teófilo y la iglesia',
        proposito: 'Relatar el crecimiento y expansión de la iglesia primitiva',
        tema: 'La obra del Espíritu Santo en la iglesia',
        bosquejo: [
          'Iglesia en Jerusalén (1–7)',
          'Expansión a Judea y Samaria (8–12)',
          'Viajes misioneros de Pablo (13–28)'
        ]
      },
      {
        libro: 'Romanos',
        autor: 'Pablo',
        fecha: '57 d.C.',
        destinatarios: 'Cristianos en Roma',
        proposito: 'Explicar claramente el evangelio y la justificación por la fe',
        tema: 'La justicia de Dios',
        bosquejo: [
          'Condenación del hombre (1–3)',
          'Justificación por fe (4–5)',
          'Santificación (6–8)',
          'Israel y el plan de Dios (9–11)',
          'Vida práctica cristiana (12–16)'
        ]
      },
      {
        libro: 'Primera de Corintios',
        autor: 'Pablo',
        fecha: '55 d.C.',
        destinatarios: 'Iglesia en Corinto',
        proposito: 'Corregir problemas doctrinales y morales',
        tema: 'Orden y santidad en la iglesia',
        bosquejo: [
          'Divisiones en la iglesia (1–4)',
          'Problemas morales y sociales (5–10)',
          'Orden en el culto (11–14)',
          'Resurrección (15)',
          'Conclusión (16)'
        ]
      },
      {
        libro: 'Segunda de Corintios',
        autor: 'Pablo',
        fecha: '56 d.C.',
        destinatarios: 'Iglesia en Corinto',
        proposito: 'Defender su apostolado y fortalecer a la iglesia',
        tema: 'El ministerio cristiano',
        bosquejo: [
          'Consuelo y ministerio (1–7)',
          'Ofrenda para los santos (8–9)',
          'Defensa apostólica (10–13)'
        ]
      },
      {
        libro: 'Gálatas',
        autor: 'Pablo',
        fecha: '48–49 d.C.',
        destinatarios: 'Iglesias de Galacia',
        proposito: 'Defender la salvación por gracia mediante la fe',
        tema: 'Libertad en Cristo',
        bosquejo: [
          'Defensa del evangelio (1–2)',
          'Doctrina de la justificación (3–4)',
          'Vida en el Espíritu (5–6)'
        ]
      },
      {
        libro: 'Efesios',
        autor: 'Pablo',
        fecha: '60–62 d.C.',
        destinatarios: 'Iglesia en Éfeso',
        proposito: 'Enseñar la posición y unidad del creyente en Cristo',
        tema: 'La iglesia, cuerpo de Cristo',
        bosquejo: [
          'Riquezas espirituales (1–3)',
          'Vida práctica cristiana (4–6)'
        ]
      },
      {
        libro: 'Filipenses',
        autor: 'Pablo',
        fecha: '61–62 d.C.',
        destinatarios: 'Iglesia en Filipos',
        proposito: 'Agradecer el apoyo recibido y exhortar al gozo',
        tema: 'Gozo en Cristo',
        bosquejo: [
          'Gozo en el ministerio (1)',
          'Gozo en la humildad (2)',
          'Gozo en la fe (3)',
          'Gozo en toda circunstancia (4)'
        ]
      },
      {
        libro: 'Colosenses',
        autor: 'Pablo',
        fecha: '60–62 d.C.',
        destinatarios: 'Iglesia en Colosas',
        proposito: 'Exaltar la supremacía de Cristo',
        tema: 'La preeminencia de Cristo',
        bosquejo: [
          'Superioridad de Cristo (1–2)',
          'Vida nueva en Cristo (3–4)'
        ]
      },
      {
        libro: 'Primera de Tesalonicenses',
        autor: 'Pablo',
        fecha: '50–51 d.C.',
        destinatarios: 'Iglesia en Tesalónica',
        proposito: 'Animar a creyentes perseguidos y enseñar sobre la segunda venida',
        tema: 'La esperanza del regreso de Cristo',
        bosquejo: [
          'Recuerdo del ministerio (1–3)',
          'Vida santa y esperanza futura (4–5)'
        ]
      },
      {
        libro: 'Segunda de Tesalonicenses',
        autor: 'Pablo',
        fecha: '51–52 d.C.',
        destinatarios: 'Iglesia en Tesalónica',
        proposito: 'Corregir errores sobre el día del Señor',
        tema: 'El día del Señor',
        bosquejo: [
          'Perseverancia en la persecución (1)',
          'Revelación del hombre de pecado (2)',
          'Exhortaciones prácticas (3)'
        ]
      },
      {
        libro: 'Primera de Timoteo',
        autor: 'Pablo',
        fecha: '62–64 d.C.',
        destinatarios: 'Timoteo',
        proposito: 'Instruir sobre organización y doctrina en la iglesia',
        tema: 'Orden en la iglesia',
        bosquejo: [
          'Sana doctrina (1)',
          'Instrucciones para la iglesia (2–3)',
          'Consejos pastorales (4–6)'
        ]
      },
      {
        libro: 'Segunda de Timoteo',
        autor: 'Pablo',
        fecha: '66–67 d.C.',
        destinatarios: 'Timoteo',
        proposito: 'Animar a permanecer fiel al ministerio',
        tema: 'Fidelidad en el servicio',
        bosquejo: [
          'Perseverar en el evangelio (1–2)',
          'Peligros de los últimos tiempos (3)',
          'Encargo final (4)'
        ]
      },
      {
        libro: 'Tito',
        autor: 'Pablo',
        fecha: '63–65 d.C.',
        destinatarios: 'Tito',
        proposito: 'Organizar las iglesias en Creta',
        tema: 'Buenas obras y sana doctrina',
        bosquejo: [
          'Líderes de la iglesia (1)',
          'Conducta cristiana (2)',
          'Buenas obras (3)'
        ]
      },
      {
        libro: 'Filemón',
        autor: 'Pablo',
        fecha: '60–62 d.C.',
        destinatarios: 'Filemón',
        proposito: 'Interceder por Onésimo',
        tema: 'Perdón y reconciliación cristiana',
        bosquejo: [
          'Saludo y gratitud (1–7)',
          'Petición por Onésimo (8–21)',
          'Despedida (22–25)'
        ]
      },
      {
        libro: 'Hebreos',
        autor: 'Desconocido',
        fecha: '65–69 d.C.',
        destinatarios: 'Cristianos judíos',
        proposito: 'Mostrar la superioridad de Cristo',
        tema: 'Cristo, superior a todo',
        bosquejo: [
          'Superioridad de Cristo (1–10)',
          'Vida de fe (11–13)'
        ]
      },
      {
        libro: 'Santiago',
        autor: 'Santiago, hermano del Señor',
        fecha: '44–49 d.C.',
        destinatarios: 'Creyentes judíos dispersos',
        proposito: 'Enseñar una fe práctica',
        tema: 'La fe demostrada por obras',
        bosquejo: [
          'Pruebas y sabiduría (1)',
          'Fe y obras (2)',
          'Dominio de la lengua (3)',
          'Vida piadosa (4–5)'
        ]
      },
      {
        libro: 'Primera de Pedro',
        autor: 'Pedro',
        fecha: '62–64 d.C.',
        destinatarios: 'Cristianos perseguidos',
        proposito: 'Animar en medio del sufrimiento',
        tema: 'Esperanza en el sufrimiento',
        bosquejo: [
          'Salvación y esperanza (1–2)',
          'Conducta cristiana (3–4)',
          'Exhortaciones finales (5)'
        ]
      },
      {
        libro: 'Segunda de Pedro',
        autor: 'Pedro',
        fecha: '64–67 d.C.',
        destinatarios: 'Creyentes en general',
        proposito: 'Advertir contra falsos maestros',
        tema: 'Conocimiento verdadero y vigilancia',
        bosquejo: [
          'Crecimiento espiritual (1)',
          'Falsos maestros (2)',
          'Día del Señor (3)'
        ]
      },
      {
        libro: 'Primera de Juan',
        autor: 'Juan',
        fecha: '85–95 d.C.',
        destinatarios: 'Iglesias de Asia Menor',
        proposito: 'Dar seguridad de salvación y combatir errores doctrinales',
        tema: 'Comunión y amor',
        bosquejo: [
          'Dios es luz (1–2)',
          'Dios es amor (3–4)',
          'Dios es vida (5)'
        ]
      },
      {
        libro: 'Segunda de Juan',
        autor: 'Juan',
        fecha: '85–95 d.C.',
        destinatarios: 'La señora elegida y su familia',
        proposito: 'Advertir contra falsos maestros',
        tema: 'Permanecer en la verdad',
        bosquejo: [
          'Caminar en amor y verdad (1–6)',
          'Advertencia contra engañadores (7–13)'
        ]
      },
      {
        libro: 'Tercera de Juan',
        autor: 'Juan',
        fecha: '85–95 d.C.',
        destinatarios: 'Gayo',
        proposito: 'Elogiar la hospitalidad y denunciar malas conductas',
        tema: 'Fidelidad y servicio cristiano',
        bosquejo: [
          'Elogio a Gayo (1–8)',
          'Problema con Diótrefes (9–12)',
          'Despedida (13–15)'
        ]
      },
      {
        libro: 'Judas',
        autor: 'Judas, hermano de Jacobo',
        fecha: '65–80 d.C.',
        destinatarios: 'Creyentes en general',
        proposito: 'Exhortar a contender por la fe',
        tema: 'Defensa de la fe verdadera',
        bosquejo: [
          'Llamado a contender por la fe (1–4)',
          'Juicio contra falsos maestros (5–16)',
          'Exhortaciones finales (17–25)'
        ]
      },
      {
        libro: 'Apocalipsis',
        autor: 'Juan',
        fecha: '95–96 d.C.',
        destinatarios: 'Las siete iglesias de Asia',
        proposito: 'Revelar el triunfo final de Cristo',
        tema: 'La victoria final de Jesucristo',
        bosquejo: [
          'Cristo glorificado (1)',
          'Mensajes a las iglesias (2–3)',
          'Juicios y tribulación (4–18)',
          'Segunda venida y reino (19–20)',
          'Cielo nuevo y tierra nueva (21–22)'
        ]
      }
    ]
  }

}
