<?php

namespace App\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\HttpFoundation\JsonResponse;

use Doctrine\ORM\EntityManagerInterface;
use App\Entity\Users;
use App\Form\UsersType;
use Symfony\Component\HttpFoundation\Request;

use Spipu\Html2Pdf\Html2Pdf;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;

// Importamos PhpOffice
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill; 

class UserController extends AbstractController
{

    public function getUsers(EntityManagerInterface $em){

      $listUsers= $em->getRepository(Users::class)->findBy([],['name' => 'ASC']);
      return $this->render('user/users.html.twig',[
        'listUsers'=> $listUsers
      ]);
    }

    public function createUser(Request $request, EntityManagerInterface $em){

      $users = new Users();
      $form_users = $this->createForm(UsersType::class, $users);
      $form_users->handleRequest($request);

      if($form_users->isSubmitted() && $form_users->isValid()){
        $users->setStatus(1);
        $em->persist($users);
        $em->flush();
        return $this->redirectToRoute('getUsers');
      }

      return $this->render('user/user_create.html.twig',[
        'form_users' => $form_users->createView()
      ]);

    }
    

    public function updateUser(Request $request, $id, EntityManagerInterface $em){
      $users = $em->getRepository(Users::class)->find($id);
      
      $form_users = $this->createForm(UsersType::class, $users);
      $form_users->handleRequest($request);
    
      if($form_users->isSubmitted() && $form_users->isValid()){

          $em->persist($users);
          $em->flush();
          return $this->redirectToRoute('getUsers');
        }

        return $this->render('user/user_update.html.twig',[
          'form_users' => $form_users->createView()
        ]);

    }
    

    public function deleteUser($id, EntityManagerInterface $em){
      $users = $em->getRepository(Users::class)->find($id);
      
      $users->setStatus(0);
      $em->persist($users);
      $em->flush();

      return $this->redirectToRoute('getUsers');
    }

    /**
     * PDF
     */
    public function getPDF(Request $request, EntityManagerInterface $em){

      $listUsers= $em->getRepository(Users::class)->findBy([],['name' => 'ASC']);
      ob_start();

      $html =$this->renderView('reports/users.html.twig',[
        'listUsers'=>$listUsers
      ]);

      $html2pdf = new Html2Pdf('P', 'LETTER', 'fr', true, 'UTF-8', array('10', '10', '10', '10'));
      $html2pdf->pdf->SetDisplayMode('real');
      $html2pdf->setDefaultFont('helvetica');
      $html2pdf->writeHTML($html);

      $cadena = 'Users.pdf';
      $originales = 'ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûýýþÿŔŕ';
      $modificadas = 'aaaaaaaceeeeiiiidnoooooouuuuybsaaaaaaaceeeeiiiidnoooooouuuyybyRr';

      $cadena = utf8_decode($cadena);
      $cadena = strtr($cadena, utf8_decode($originales), $modificadas);
      $cadena = strtoupper($cadena);
     
      ob_end_clean();

      return new Response($html2pdf->Output(utf8_encode($cadena), 'D'), 200, [
          'Content-Type' => 'application/pdf;charset=UTF-8'
      ]);

    }

    /**
     * EXCEL
     */
    public function getExcel(EntityManagerInterface $entityManager){

      $listUsers = $entityManager->getRepository(Users::class)->findBy([], ['name' => 'ASC']);

      $spreadsheet = new Spreadsheet();

      $sheet =  $spreadsheet->getActiveSheet();

      $sheet->setCellValue("A1","#");
      $sheet->setCellValue("B1","Nombre completo");
      $sheet->setCellValue("C1","Correo electronico");

      $style=[
              'font'=>[
                  'bold'=> true,
                  'color'=> ['rgb' =>'FFFFFF'],
                  'size'=>12,
                  'name' => 'Century Gothic'
              ],
              'alignment' =>[
                  'vertical'=> Alignment::VERTICAL_CENTER,
                  'horizontal'=> Alignment::HORIZONTAL_CENTER
              ]
          ];

      $sheet->getStyle("A1:C1")->getFill()->setFillType(Fill::FILL_SOLID);
      $sheet->getStyle("A1:C1")->getFill()->getStartColor()->setRGB("012756");

      for ($i = 0; $i < count($listUsers); $i++) {
          $counter = $i + 2;
          $sheet->setCellValue("A" . $counter, $i + 1);
          $sheet->getStyle("A" . $counter)->getFill()->setFillType(Fill::FILL_SOLID);
          $sheet->getStyle("A" . $counter)->getFill()->getStartColor()->setRGB("012756");
          $sheet->getStyle("A" . $counter)->applyFromArray($style);

          $sheet->setCellValue("B" . $counter, $listUsers[$i]->getName() . " " . $listUsers[$i]->getLastname());
          $sheet->setCellValue("C" . $counter, $listUsers[$i]->getEmail());
      }

      $sheet->getStyle('A1:C1')->applyFromArray($style);

      $sheet->setTitle("Usuarios");

      $sheet->getColumnDimension("B")->setWidth(30);
      $sheet->getColumnDimension("C")->setWidth(30);

      // new element xlsx
      $writer = new Xlsx($spreadsheet);

      // file name and temporal file
      $actualDate =(new \DateTime())->format('d-m-Y');
      $fileName = $actualDate . '.xlsx';
      $temp_file = tempnam(sys_get_temp_dir(), $fileName);
      
      // save temporal file and return it
      $writer->save($temp_file);
      return $this->file($temp_file, $fileName, ResponseHeaderBag::DISPOSITION_INLINE);

   }
}
