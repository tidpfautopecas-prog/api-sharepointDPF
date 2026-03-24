import jsPDF from 'jspdf';
import type { Ticket } from '../pages/laudos/page';
import { supabase } from './supabaseClient';

const delay = (ms: number) => new Promise(res => setTimeout(res, ms));
const STORAGE_BUCKET = 'fotos-laudos';

async function fetchImageAsBase64(path: string): Promise<string> {
    try {
        const { data: blob, error: downloadError } = await supabase
            .storage
            .from(STORAGE_BUCKET)
            .download(path);

        if (downloadError) throw downloadError;

        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result as string);
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        });
    } catch (error) {
        return '';
    }
}

export class PDFGenerator {
  private static instance: PDFGenerator;
  private toastCallback: ((type: 'success' | 'error' | 'info' | 'warning', title: string, message?: string) => void) | null = null;
  private apiUrl: string = "https://api-sharepointdpf.onrender.com";

  public static getInstance(): PDFGenerator {
    if (!PDFGenerator.instance) {
      PDFGenerator.instance = new PDFGenerator();
    }
    return PDFGenerator.instance;
  }

  setApiUrl = (url: string) => {
    this.apiUrl = url;
  }
  
  setToastCallback = (callback: (type: 'success' | 'error' | 'info' | 'warning', title: string, message?: string) => void) => {
    this.toastCallback = callback;
  }

  private showToast = (type: 'success' | 'error' | 'info' | 'warning', title: string, message?: string) => {
    if (this.toastCallback) {
        this.toastCallback(type, title, message);
    }
  }

  setSharePointConfig = (siteUrl: string, libraryName: string) => {
    const config = { siteUrl, libraryName };
    localStorage.setItem('sharePointConfig', JSON.stringify(config));
  }

  getSharePointConfig = (): { siteUrl: string, libraryName: string } | null => {
    const config = localStorage.getItem('sharePointConfig');
    return config ? JSON.parse(config) : null;
  }

  setDefaultNetworkPath = (path: string) => {
    localStorage.setItem('defaultNetworkPath', path);
  }

  getDefaultNetworkPath = (): string | null => {
    return localStorage.getItem('defaultNetworkPath');
  }

  syncSmartTicket = async (ticket: Ticket): Promise<string> => {
    try {
        const cleanTicket = ticket.numero.replace(/[^a-zA-Z0-9-]/g, '');
        const statusRes = await fetch(`${this.apiUrl}/check-status/${cleanTicket}`);
        
        let needsPdf = true;

        if (statusRes.ok) {
            const status = await statusRes.json();
            needsPdf = !status.existsInPdf;
        }

        if (!needsPdf) {
            return 'Ignorado (Já existe)';
        }

        const actions = [];

        if (needsPdf) {
            await this.generateAndUploadOnlyPdf(ticket);
            actions.push('PDF');
        }

        return `Sincronizado: ${actions.join(' e ')}`;

    } catch (error: any) {
        throw error;
    }
  }

  deletePDFByTicketNumber = async (ticketNumber: string): Promise<void> => {
    try {
        const cleanTicketNumber = ticketNumber.replace(/[^a-zA-Z0-9-]/g, '');
        const response = await fetch(`${this.apiUrl}/delete-pdf-by-ticket-number/${cleanTicketNumber}`, {
            method: 'DELETE',
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Falha na API: Status ${response.status}. Resposta: ${errorText}`);
        }
    } catch (error) {
        this.showToast('error', 'Erro ao remover PDF', 'O ficheiro pode não existir no SharePoint ou a API está offline.');
    }
  }

  deleteListDataByTicketNumber = async (ticketNumber: string): Promise<void> => {
    try {
        const cleanTicketNumber = ticketNumber.replace(/[^a-zA-Z0-9-]/g, '');
        const response = await fetch(`${this.apiUrl}/delete-list-data-by-ticket-number/${cleanTicketNumber}`, {
            method: 'DELETE',
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Falha na API: Status ${response.status}. Resposta: ${errorText}`);
        }
    } catch (error) {
        this.showToast('error', 'Erro ao remover itens da lista', 'A API pode estar offline ou o SharePoint negou o acesso.');
    }
  }

  clearSharePointList = async (): Promise<void> => {
    try {
      const response = await fetch(`${this.apiUrl}/clear-list`, {
        method: 'DELETE',
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Falha na API: Status ${response.status}. Resposta: ${errorText}`);
      }
      
      this.showToast('success', 'Lista limpa com sucesso');
    } catch (error) {
      this.showToast('error', 'Erro ao limpar lista', 'A API pode estar offline ou o SharePoint negou o acesso.');
      throw error;
    }
  }

  saveReportToSharePoint = async (pdf: jsPDF, fileName: string): Promise<void> => {
    const fileBase64 = pdf.output('datauristring').split(',')[1];
    await this.uploadWithRetries('upload-pdf', { fileName, fileBase64: fileBase64, isReport: true });
  }

  generateTicketPDF = async (ticket: Ticket, userName: string): Promise<void> => {
    try {
        const pdf = new jsPDF('p', 'mm', 'a4');
        const w = pdf.internal.pageSize.getWidth();
        const h = pdf.internal.pageSize.getHeight();
        let y = 15;

        this.addHeader(pdf, w, y);
        y += 20;
        y = this.addTicketInfo(pdf, ticket, y, w);
        y += 10;
        if (ticket.itens && ticket.itens.length > 0) {
            y = await this.addItemsList(pdf, ticket.itens, y, w, h);
        }
        this.addFooter(pdf, w, h, userName);
        
        const fileName = this.generateFileName(ticket.numero);
        const fileBase64 = pdf.output('datauristring').split(',')[1];

        await this.uploadWithRetries('upload-pdf', {
            fileName: fileName,
            fileBase64: fileBase64,
            ticketNumber: ticket.numero,
            ticketTitle: ticket.titulo,
            isReport: false,
        });

        // Gerando a Data e Hora no formato limpo para o SharePoint exibir por extenso
        const agora = new Date();
        const dataGeracaoFormatada = `${agora.toLocaleDateString('pt-BR')} ${agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}`;

        const listData = ticket.itens ? ticket.itens.map(item => {
            const rowData: any = {
                Title: `Laudo ${ticket.numero} - Item ${item.numeroItem} (Criado por: ${userName})`,
                ticketNumber: ticket.numero,
                nomeCliente: ticket.nomeCliente || '',
                item: String(item.numeroItem),
                qtde: String(item.quantidade || ''),
                motivo: item.motivo ? item.motivo.join(' / ') : '',
                origemDefeito: item.origemDefeito || '',
                disposicao: item.disposicao || '',
                disposicaoPecas: item.disposicaoPecas || '',
                dataGeracao: dataGeracaoFormatada // Envia a data bonita para a coluna Data de Geração!
            };

            if (item.fotos) {
                item.fotos.forEach((fotoPath, fIdx) => {
                    if (fIdx < 10) {
                        const { data } = supabase.storage.from(STORAGE_BUCKET).getPublicUrl(fotoPath);
                        rowData[`foto${fIdx + 1}`] = data.publicUrl;
                    }
                });
            }
            return rowData;
        }) : [];

        if (listData.length > 0) {
            await this.uploadWithRetries('upload-list-data', { listData });
        }
        
        this.showToast('success', 'Salvo no Share Point');

    } catch (apiError: any) {
        this.showToast('error', 'Erro ao salvar no Share Point. Tentando download local.');
        
        try {
            const fallbackPdf = new jsPDF('p', 'mm', 'a4');
            const w = fallbackPdf.internal.pageSize.getWidth();
            const h = fallbackPdf.internal.pageSize.getHeight();
            let currentY = 15;
            this.addHeader(fallbackPdf, w, currentY);
            currentY += 20;
            currentY = this.addTicketInfo(fallbackPdf, ticket, currentY, w);
            currentY += 10;
            if (ticket.itens && ticket.itens.length > 0) {
                await this.addItemsList(fallbackPdf, ticket.itens, currentY, w, h);
            }
            this.addFooter(fallbackPdf, w, h, userName);
            
            const fileName = this.generateFileName(ticket.numero);
            fallbackPdf.save(fileName);
        } catch (localPdfError) {
        }
    }
  }

  private generateAndUploadOnlyPdf = async (ticket: Ticket) => {
      const pdf = new jsPDF('p', 'mm', 'a4');
      const w = pdf.internal.pageSize.getWidth();
      const h = pdf.internal.pageSize.getHeight();
      let y = 15;
      this.addHeader(pdf, w, y); 
      y += 20;
      y = this.addTicketInfo(pdf, ticket, y, w); 
      y += 10;
      if (ticket.itens && ticket.itens.length > 0) {
          y = await this.addItemsList(pdf, ticket.itens, y, w, h);
      }
      this.addFooter(pdf, w, h, ticket.responsavel);
      
      const fileName = this.generateFileName(ticket.numero);
      const fileBase64 = pdf.output('datauristring').split(',')[1];
      
      await this.uploadWithRetries('upload-pdf', {
          fileName: fileName, 
          fileBase64: fileBase64, 
          ticketNumber: ticket.numero, 
          ticketTitle: ticket.titulo, 
          isReport: false
      });

      const criadorNome = (ticket as any).responsavel_nome || ticket.responsavel || 'Sistema';
      
      const agora = new Date();
      const dataGeracaoFormatada = `${agora.toLocaleDateString('pt-BR')} ${agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}`;


      const listData = ticket.itens ? ticket.itens.map(item => {
          const rowData: any = {
              Title: `Laudo ${ticket.numero} - Item ${item.numeroItem} (Criado por: ${criadorNome})`,
              ticketNumber: ticket.numero,
              nomeCliente: ticket.nomeCliente || '',
              item: String(item.numeroItem),
              qtde: String(item.quantidade || ''),
              motivo: item.motivo ? item.motivo.join(' / ') : '',
              origemDefeito: item.origemDefeito || '',
              disposicao: item.disposicao || '',
              disposicaoPecas: item.disposicaoPecas || '',
              dataGeracao: dataGeracaoFormatada
          };

          if (item.fotos) {
              item.fotos.forEach((fotoPath, fIdx) => {
                  if (fIdx < 10) {
                      const { data } = supabase.storage.from(STORAGE_BUCKET).getPublicUrl(fotoPath);
                      rowData[`foto${fIdx + 1}`] = data.publicUrl;
                  }
              });
          }
          return rowData;
      }) : [];

      if (listData.length > 0) {
          await this.uploadWithRetries('upload-list-data', { listData });
      }
  }

  private uploadWithRetries = async (endpoint: string, body: object, retries = 3, initialDelay = 2000) => {
    let attempt = 0;
    while (attempt < retries) {
      try {
        const response = await fetch(`${this.apiUrl}/${endpoint}`, {
          method: 'POST',
          headers: { 
            'Content-Type': 'application/json', 
            'Accept': 'application/json' 
          },
          body: JSON.stringify(body),
        });

        if (response.ok) {
          const result = await response.json();
          if (result.success) return result;
        }
        throw new Error(`Falha na API: Status ${response.status}`);
      } catch (error) {
        attempt++;
        if (attempt >= retries) throw error;
        const delayTime = initialDelay * Math.pow(2, attempt - 1);
        await delay(delayTime);
      }
    }
  }
  
  private generateFileName = (ticketNumero: string): string => {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const cleanTicketNumber = ticketNumero.replace(/[^a-zA-Z0-9-]/g, ''); 
    return `Laudo - ${cleanTicketNumber}-${year}${month}${day} ${hours}${minutes}.pdf`;
  }

  private addHeader = (pdf: jsPDF, pageWidth: number, y: number): void => {
    try {
      const logoUrl = 'https://i.postimg.cc/bJ3kwSbw/DPF.png';
      pdf.addImage(logoUrl, 'PNG', 21, y - 5, 23, 12);
    } catch (error) {
      pdf.setFontSize(10).setFont('helvetica', 'bold').setTextColor(0, 51, 102);
      pdf.text('DPF', 15, y + 2);
    }
    pdf.setFontSize(16).setFont('helvetica', 'bold').setTextColor(0, 0, 0);
    pdf.text('Laudo Técnico de Garantia', pageWidth / 2, y, { align: 'center' });
    pdf.setDrawColor(220, 220, 220).setLineWidth(0.5);
    pdf.line(15, y + 10, pageWidth - 15, y + 10);
  }

  private addField = (pdf: jsPDF, y: number, label: string, value: string, pageWidth: number): number => {
    pdf.setFontSize(10);
    pdf.setFont('helvetica', 'bold').setTextColor(50, 50, 50);
    pdf.text(label, 15, y);
    const labelWidth = pdf.getStringUnitWidth(label) * pdf.getFontSize() / pdf.internal.scaleFactor;
    const valueX = 15 + labelWidth + 2;
    pdf.setFont('helvetica', 'normal').setTextColor(80, 80, 80);
    const valueAvailableWidth = pageWidth - valueX - 15;
    const valueLines = pdf.splitTextToSize(value || 'Não informado', valueAvailableWidth);
    pdf.text(valueLines, valueX, y);
    return y + (valueLines.length * 5) + 3;
  }

  private addTicketInfo = (pdf: jsPDF, ticket: Ticket, y: number, pageWidth: number): number => {
    pdf.setFontSize(14).setFont('helvetica', 'bold').setTextColor(0, 0, 0);
    pdf.text('INFORMAÇÕES DO LAUDO', 15, y);
    y += 8;
    y = this.addField(pdf, y, 'Número do Ticket:', ticket.numero, pageWidth);
    y = this.addField(pdf, y, 'Nome do Cliente:', ticket.nomeCliente || 'Não informado', pageWidth);
    return y;
  }

  private addItemsList = async (pdf: jsPDF, itens: Ticket['itens'], y: number, pageWidth: number, pageHeight: number): Promise<number> => {
    pdf.setFontSize(14).setFont('helvetica', 'bold').setTextColor(0, 0, 0);
    pdf.text('ITENS DO LAUDO', 15, y);
    y += 10;
    for (let i = 0; i < itens.length; i++) {
      const item = itens[i];
      let itemStartY = y;
      const estimatedHeight = 80 + (item.fotos && item.fotos.length > 0 ? 60 : 0);
      if (itemStartY + estimatedHeight > pageHeight - 20) {
        pdf.addPage();
        y = 20;
        itemStartY = y;
      }
      pdf.setDrawColor(0, 51, 102).setFillColor(240, 248, 255);
      pdf.roundedRect(15, itemStartY - 5, pageWidth - 30, 10, 2, 2, 'FD');
      pdf.setFontSize(12).setFont('helvetica', 'bold').setTextColor(0, 31, 63);
      pdf.text(`ITEM: ${item.numeroItem}`, 20, itemStartY);
      pdf.setFontSize(9).setFont('helvetica', 'normal').setTextColor(100, 100, 100);
      pdf.text(`Quantidade: ${item.quantidade}`, pageWidth - 20, itemStartY, { align: 'right' });
      y += 10;
      const motivosString = item.motivo.join('\n'); 
      
      const origemFormatada = item.origemDefeito 
        ? item.origemDefeito.replace(/(\d+):\s*-\s*/, ' $1 - ')
        : item.origemDefeito;

      const disposicaoPecasFormatada = item.disposicaoPecas
        ? item.disposicaoPecas.replace(/(\d+)-\s*/, ' $1 - ')
        : item.disposicaoPecas;

      y = this.addField(pdf, y, 'Motivo:', motivosString, pageWidth); 
      y = this.addField(pdf, y, 'Origem do Defeito:', origemFormatada, pageWidth);
      y = this.addField(pdf, y, 'Disposição:', item.disposicao, pageWidth);
      y = this.addField(pdf, y, 'Disposição das Peças:', disposicaoPecasFormatada, pageWidth);
      
      if (item.anotacoes) {
        y = this.addField(pdf, y, 'Anotações:', item.anotacoes, pageWidth);
      }
      y += 3;
      if (item.fotos && item.fotos.length > 0) {
        const photoSize = 45;
        const gap = 4;
        const photosPerRow = 4;
        let photoX = 15;
        if (y + photoSize > pageHeight - 20) {
          pdf.addPage();
          y = 20;
        }
        pdf.setFontSize(10).setFont('helvetica', 'bold').setTextColor(50, 50, 50);
        pdf.text('Fotos:', 15, y);
        y += 7;
        
        const imageBase64Strings = await Promise.all(
            item.fotos.map(path => fetchImageAsBase64(path))
        );

        for(let p = 0; p < imageBase64Strings.length; p++) {
          const fotoBase64 = imageBase64Strings[p];
          if (p > 0 && p % photosPerRow === 0) {
            y += photoSize + gap;
            photoX = 15;
            if (y + photoSize > pageHeight - 20) {
              pdf.addPage();
              y = 20;
            }
          }
          
          if (fotoBase64) { 
            try {
              pdf.addImage(fotoBase64, 'JPEG', photoX, y, photoSize, photoSize);
              pdf.setDrawColor(200, 200, 200).rect(photoX, y, photoSize, photoSize);
            } catch (e) {
              pdf.setFontSize(8).setTextColor(150, 0, 0);
              pdf.text('[img erro]', photoX + photoSize / 2, y + photoSize / 2, { align: 'center' });
            }
          } else {
            pdf.setDrawColor(150, 0, 0).rect(photoX, y, photoSize, photoSize);
            pdf.setFontSize(8).setTextColor(150, 0, 0);
            pdf.text('[img erro]', photoX + photoSize / 2, y + photoSize / 2, { align: 'center' });
          }
          photoX += photoSize + gap;
        }
        y += photoSize + 10;
      }
      if (i < itens.length - 1) {
        pdf.setDrawColor(220, 220, 220).setLineWidth(0.3);
        pdf.line(15, y, pageWidth - 15, y);
        y += 10;
      }
    }
    return y;
  }

  private addFooter = (pdf: jsPDF, pageWidth: number, pageHeight: number, userName: string): void => {
    const totalPages = (pdf as any).internal.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
      pdf.setPage(i);
      pdf.setDrawColor(220, 220, 220).setLineWidth(0.5);
      pdf.line(15, pageHeight - 15, pageWidth - 15, pageHeight - 15);
      pdf.setFontSize(8).setFont('helvetica', 'normal').setTextColor(100, 100, 100);
      pdf.text('DPF AUTO PEÇAS LTDA - Sistema de Laudos', 15, pageHeight - 10);
      pdf.text(`Página ${i} de ${totalPages}`, pageWidth / 2, pageHeight - 10, { align: 'center' });
      const dataGeracao = new Date().toLocaleString('pt-BR');
      pdf.text(`Gerado por ${userName} em ${dataGeracao}`, pageWidth - 15, pageHeight - 10, { align: 'right' });
    }
  }
}

export const pdfGenerator = PDFGenerator.getInstance();
